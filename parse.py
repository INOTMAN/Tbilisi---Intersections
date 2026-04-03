#!/usr/bin/env python3
"""Parse vakee.zip Excel files into data.js for the intersections map.

Three file formats exist:
  VAK2/CHU2: row0 has 'General' in col0; code at col2, name col3, X(lng) col4, Y(lat) col5
             TL section starts with 'Traffic Lights' in col0; TL name at col1
  VAK3:      row0 has 'Number_' in col0 (no 'General'); code at col1, name col2, X col3, Y col4
             TL section starts with 'Name','Diameter' in col0,col1; TL name at col0
"""
import zipfile, re, json, os, sys
from xml.etree import ElementTree as ET

ZIP_PATH = os.path.join(os.path.dirname(__file__), 'vakee.zip')
OUT_PATH  = os.path.join(os.path.dirname(__file__), 'data.js')
CHU_PATH  = '/mnt/c/Users/Ancho Jokhadze/Desktop/CHU2-013.xlsx'

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

# Intersection codes: VAK2-009, VAK3-054, CHU2-013 etc.
INTER_CODE_RE = re.compile(r'^[A-Z]{2,4}\d-\d{3,4}$')
# Traffic light names: CP121-1, SP153-6, PP, SP, CP etc.
TL_NAME_RE = re.compile(r'^(CP|SP|PP)[0-9\-]*$', re.IGNORECASE)

def is_inter_code(s):
    return bool(INTER_CODE_RE.match(str(s).strip()))

def is_tl_name(s):
    return bool(TL_NAME_RE.match(str(s).strip()))

def col_to_idx(col):
    idx = 0
    for ch in col:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def safe(row, i):
    if i < len(row):
        v = str(row[i]).strip()
        # Strip trailing .0 from numbers like "22.0" -> "22" only if it's a whole number
        if re.match(r'^\d+\.0$', v):
            return v[:-2]
        return v
    return ''

def parse_float(s):
    try:
        return round(float(s), 6)
    except:
        return 0.0

def load_xlsx_zip(zf, name):
    """Load xlsx from inner ZipFile entry."""
    try:
        inner = zipfile.ZipFile(zf.open(name))
        return _parse_inner(inner)
    except Exception as e:
        print(f"  Cannot open {name}: {e}")
        return []

def load_xlsx_path(path):
    """Load xlsx from filesystem path."""
    inner = zipfile.ZipFile(path)
    return _parse_inner(inner)

def _parse_inner(inner):
    shared = []
    if 'xl/sharedStrings.xml' in inner.namelist():
        tree = ET.parse(inner.open('xl/sharedStrings.xml'))
        for si in tree.getroot().findall(f'.//{NS}si'):
            parts = si.findall(f'.//{NS}t')
            shared.append(''.join((p.text or '') for p in parts))
    if 'xl/worksheets/sheet1.xml' not in inner.namelist():
        return []
    tree = ET.parse(inner.open('xl/worksheets/sheet1.xml'))
    root = tree.getroot()
    rows_out = []
    for row in root.findall(f'.//{NS}row'):
        cells = {}
        for c in row.findall(f'{NS}c'):
            r = c.get('r', '')
            m = re.match(r'([A-Z]+)', r)
            if not m:
                continue
            ci = col_to_idx(m.group(1))
            t = c.get('t', '')
            v_el = c.find(f'{NS}v')
            val = ''
            if v_el is not None and v_el.text:
                if t == 's':
                    try:
                        val = shared[int(v_el.text)]
                    except:
                        val = v_el.text
                else:
                    val = v_el.text
            cells[ci] = val
        if cells:
            max_idx = max(cells.keys())
            rows_out.append([cells.get(i, '') for i in range(max_idx + 1)])
    return rows_out


def detect_format(rows):
    """Return 'vak2' or 'vak3' based on first row structure."""
    if not rows:
        return 'vak2'
    row0 = [str(c).lower().strip() for c in rows[0]]
    if row0 and 'general' in row0[0]:
        return 'vak2'
    return 'vak3'


def parse_intersections_meta(rows, fmt):
    """Extract intersection metadata rows, including the intersection NUMBER used in TL names."""
    intersections = []
    if fmt == 'vak2':
        # col1=number, col2=code, col3=name, col4=X(lng), col5=Y(lat), col6=status, col7=date, col8=persons
        for row in rows:
            code = safe(row, 2)
            if is_inter_code(code):
                try:
                    number = int(float(safe(row, 1)))
                except:
                    number = None
                intersections.append({
                    'number': number,
                    'code': code,
                    'name': safe(row, 3),
                    'lng': parse_float(safe(row, 4)),
                    'lat': parse_float(safe(row, 5)),
                    'status': safe(row, 6),
                    'date': safe(row, 7),
                    'persons': safe(row, 8),
                })
    else:  # vak3
        # col0=number, col1=code, col2=name, col3=X(lng), col4=Y(lat), col5=status, col6=date, col7=persons
        for row in rows:
            code = safe(row, 1)
            if is_inter_code(code):
                try:
                    number = int(float(safe(row, 0)))
                except:
                    number = None
                intersections.append({
                    'number': number,
                    'code': code,
                    'name': safe(row, 2),
                    'lng': parse_float(safe(row, 3)),
                    'lat': parse_float(safe(row, 4)),
                    'status': safe(row, 5),
                    'date': safe(row, 6),
                    'persons': safe(row, 7),
                })
    return intersections


def tl_number(tl_name):
    """Extract intersection number from TL name. CP131-1 → 131, SP153-6 → 153. Returns None if no number."""
    m = re.match(r'^[A-Z]{2,3}P?(\d+)', tl_name.strip())
    if m:
        return int(m.group(1))
    return None


def find_approach_section(rows, fmt):
    """Find the index of the approach data header row.
    Handles typos like 'Appraches', 'Apporach', 'FID'.
    """
    for i, row in enumerate(rows):
        c0 = str(row[0]).strip().lower() if row else ''
        c1 = str(row[1]).strip().lower() if len(row) > 1 else ''
        # Typo-tolerant approach header detection
        if 'appr' in c0 or c0 == 'fid':
            return i
        if 'direction' in c1 and c0 not in ('', 'general', 'number_'):
            return i
    return -1


def parse_approaches(rows, hdr_idx, fmt):
    """Parse approach rows starting after hdr_idx."""
    result = []
    num_col = 0  # approach number column
    for row in rows[hdr_idx + 1:]:
        num_str = safe(row, num_col)
        if not num_str:
            continue
        # Check for end of section
        c0_low = str(row[0]).lower().strip()
        if any(k in c0_low for k in ['traffic', 'light', 'name', 'file', 'cad']):
            break
        try:
            num = int(float(num_str))
        except:
            continue
        # VAK2: col0=num, col1=direction, col2=width, ...
        # VAK3: col0=num, col1=direction, col2=width, ...  (same structure)
        def av(i): return safe(row, i)
        result.append({
            'num': num,
            'direction': av(1),
            'width': av(2),
            'traffic': av(3),
            'turning': av(4),
            'buslane': av(5),
            'bikelane': av(6),
            'bikeway': av(7),
            'parking': av(8),
            'buffer': av(9),
            'tc': av(10),
            'power': av(11),
            'internet': av(12),
            'powerDem': av(13),
            'intDem': av(14),
            'railing': av(15),
            'busStop': av(16),
            'shelter': av(17),
            'adPanel': av(18),
            'panelDim': av(19),
            'notes': av(20),
        })
    return result


def find_tl_section(rows, fmt):
    """Find the index of the TL data header row.
    VAK2: row with col0='Traffic Lights' — data rows follow (name at col1)
    VAK3: row with col0='Name', col1='Diameter' — data rows follow (name at col0)
    """
    for i, row in enumerate(rows):
        c0 = str(row[0]).strip().lower()
        c1 = str(row[1]).strip().lower() if len(row) > 1 else ''
        if fmt == 'vak2' and 'traffic' in c0 and 'light' in c0:
            # Next row should be the column header; return index of column header
            for j in range(i, min(i + 3, len(rows))):
                r0 = str(rows[j][0]).strip().lower()
                r1 = str(rows[j][1]).strip().lower() if len(rows[j]) > 1 else ''
                if r1 == 'name' or r0 == 'traffic lights':
                    return j  # actual header row with Name column
            return i
        if fmt == 'vak3' and c0 == 'name' and 'diam' in c1:
            return i
    return -1


def parse_tl(rows, hdr_idx, fmt):
    """Parse traffic light rows after hdr_idx.
    VAK2: name at col1, data cols 2-21
    VAK3: name at col0, data cols 1-20
    """
    result = []
    for row in rows[hdr_idx + 1:]:
        if not any(str(c).strip() for c in row):
            continue
        c0 = str(row[0]).strip()
        c1 = str(row[1]).strip() if len(row) > 1 else ''

        # Detect name and offset
        if fmt == 'vak3':
            if not is_tl_name(c0):
                break
            name = c0
            def tv(i): return safe(row, i + 1)
        else:  # vak2 / CHU2
            if is_tl_name(c1):
                name = c1
                def tv(i): return safe(row, i + 2)
            elif is_tl_name(c0):
                name = c0
                def tv(i): return safe(row, i + 1)
            else:
                # Only stop on a real section header, not on empty/note-only rows
                c0_low = c0.lower()
                if c0_low and c0_low not in ('', 'files', 'cad'):
                    break  # clear section header (e.g. "Files", "CAD")
                # Empty row or note-only row — skip and continue
                continue

        result.append({
            'name': name,
            'diameter': tv(0),
            'height': tv(1),
            'length': tv(2),
            'traffic': tv(3),
            'pedestrian': tv(4),
            'cycling': tv(5),
            'singleR': tv(6),
            'singleY': tv(7),
            'singleG': tv(8),
            'countdown': tv(9),
            'forward': tv(10),
            'right': tv(11),
            'left': tv(12),
            'bus': tv(13),
            'camera': tv(14),
            'detector': tv(15),
            'acoustic': tv(16),
            'button': tv(17),
            'sunshade': tv(18),
            'note': tv(19),
        })
    return result


def parse_file(rows):
    """Extract all intersections from a parsed row list."""
    if not rows:
        return []

    fmt = detect_format(rows)
    meta_list = parse_intersections_meta(rows, fmt)

    if not meta_list:
        return []

    app_hdr = find_approach_section(rows, fmt)
    tl_hdr  = find_tl_section(rows, fmt)

    approaches = parse_approaches(rows, app_hdr, fmt) if app_hdr >= 0 else []
    all_tls    = parse_tl(rows, tl_hdr, fmt)          if tl_hdr >= 0 else []

    # Split TLs by intersection number embedded in TL name (e.g. CP131-1 → 131 → VAK2-018)
    # Build map: intersection number → meta index
    num_to_idx = {}
    for i, meta in enumerate(meta_list):
        if meta['number'] is not None:
            num_to_idx[meta['number']] = i

    if len(meta_list) > 1 and num_to_idx:
        tl_per_inter = [[] for _ in meta_list]
        unmatched_tls = []
        for tl in all_tls:
            n = tl_number(tl['name'])
            if n is not None and n in num_to_idx:
                tl_per_inter[num_to_idx[n]].append(tl)
            else:
                unmatched_tls.append(tl)
        # Unmatched TLs (no number in name) go to all intersections
        for lst in tl_per_inter:
            lst.extend(unmatched_tls)
    else:
        tl_per_inter = [all_tls] * len(meta_list)

    intersections = []
    for i, meta in enumerate(meta_list):
        intersections.append({
            'code': meta['code'],
            'name': meta['name'],
            'lat': meta['lat'],
            'lng': meta['lng'],
            'status': meta['status'],
            'date': meta['date'],
            'persons': meta['persons'],
            'approaches': approaches,
            'trafficLights': tl_per_inter[i],
            'files': [],
        })

    return intersections


def main():
    if not os.path.exists(ZIP_PATH):
        print(f"ERROR: {ZIP_PATH} not found")
        sys.exit(1)

    outer = zipfile.ZipFile(ZIP_PATH)
    xlsx_files = sorted(n for n in outer.namelist() if n.lower().endswith('.xlsx'))
    print(f"Found {len(xlsx_files)} xlsx files in zip")

    all_intersections = []
    seen_codes = set()

    def add(items):
        for item in items:
            code = item['code']
            if code and code not in seen_codes:
                seen_codes.add(code)
                all_intersections.append(item)
                print(f"  + {code}: {len(item['trafficLights'])} lights, "
                      f"{len(item['approaches'])} approaches")
            else:
                print(f"  (skip dup {code})")

    # CHU2-013 from Desktop
    if os.path.exists(CHU_PATH):
        print(f"Loading {CHU_PATH}")
        add(parse_file(load_xlsx_path(CHU_PATH)))

    for xlsx_name in xlsx_files:
        print(f"Loading {xlsx_name}")
        rows = load_xlsx_zip(outer, xlsx_name)
        add(parse_file(rows))

    all_intersections.sort(key=lambda x: x['code'])
    print(f"\nTotal intersections: {len(all_intersections)}")

    lights_ok = sum(1 for x in all_intersections if x['trafficLights'])
    print(f"With traffic lights: {lights_ok}")

    # Write data.js
    lines = ['const INTERSECTIONS = [']
    for item in all_intersections:
        lines.append('  ' + json.dumps(item, ensure_ascii=False) + ',')
    lines.append('];')

    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print(f"Written to {OUT_PATH}")


if __name__ == '__main__':
    main()
