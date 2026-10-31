"""PQM summary dashboard. Run: streamlit run pq_dash.py"""
from io import BytesIO
from html import escape
from numbers import Real
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
SHEETS = {'Zone': 'Zone Summary', 'Region(AVP)': 'Region(AVP) Summary', 'DM': 'DM Summary', 'PQM': 'PQM Summary'}
IDENTITY = {'Zone': ['Zone','Zonal Name','PQ RM Name'], 'Region(AVP)': ['Zone','Region AVP','Region AVP Name','PQ DM Name'], 'DM': ['RM','DM Name'], 'PQM': ['Zone','Region AVP','State','RM Name','DM Name','PQM Name','PQM EMP ID','Branch']}

def colnum(ref):
    result = 0
    for c in re.match(r'[A-Z]+', ref)[0]:
        result = result * 26 + ord(c) - 64
    return result


def coord(ref):
    return int(re.search(r'\d+', ref)[0]), colnum(ref)


def load_report(raw):
    """Read only summary XML and cached results; never recompute business logic."""
    result = {}
    with zipfile.ZipFile(BytesIO(raw)) as z:
        if sum(i.file_size for i in z.infolist()) > 800_000_000:
            raise ValueError('Workbook is too large. Please upload a smaller report.')
        strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            strings = [''.join(t.text or '' for t in si.findall('.//s:t', NS))
                       for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        rels = {r.attrib['Id']: r.attrib['Target'] for r in
                ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = {}
        for sh in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS):
            rid = sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target = rels[rid]
            sheets[sh.attrib['name']] = (target.lstrip('/') if target.startswith('/')
                                        else posixpath.normpath('xl/' + target))
        for level, name in SHEETS.items():
            if name not in sheets:
                raise ValueError(f'Missing sheet: {name}. Upload the Combined PQM Report .xlsx.')
            root = ET.fromstring(z.read(sheets[name]))
            cells, missing = {}, 0
            width = max((colnum(c.attrib['r']) for c in root.findall('s:sheetData/s:row/s:c', NS) if int(re.search(r'\d+', c.attrib['r'])[0]) <= 3), default=0)
            # Ignore unrelated trailing helper columns and notes.
            width = min(width, 250)  # dynamic header width, including newly inserted columns
            if width < 10 or width > 300: raise ValueError(f'Unexpected number of columns in {name}: {width}')
            for cell in root.findall('s:sheetData/s:row/s:c', NS):
                r, c = coord(cell.attrib['r'])
                if c > width:
                    continue
                v = cell.find('s:v', NS)
                value = v.text if v is not None else None
                kind = cell.attrib.get('t')
                if kind == 's' and value is not None:
                    value = strings[int(value)]
                elif kind == 'inlineStr':
                    value = ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
                elif value is not None and kind not in ('e', 'str'):
                    try:
                        value = float(value)
                        if value.is_integer():
                            value = int(value)
                    except ValueError:
                        pass
                if cell.find('s:f', NS) is not None and value is None:
                    missing += 1
                if value is not None:
                    cells[r, c] = value
            merges = []
            for m in root.findall('s:mergeCells/s:mergeCell', NS):
                a, b = m.attrib['ref'].split(':')
                r1, c1 = coord(a); r2, c2 = coord(b)
                if r1 <= 3 and c1 <= width:
                    merges.append((r1, c1, r2, min(c2, width)))
            headers = []
            ids = len(IDENTITY[level]) if level != 'DM' else 4
            for c in range(1, width + 1):
                parts = []
                for r in range(1, 4):
                    rr, cc = r, c
                    for a, b, d, e in merges:
                        if a <= r <= d and b <= c <= e:
                            rr, cc = a, b
                            break
                    value = str(cells.get((rr, cc), '')).replace('\n', ' ').strip()
                    if value and not value.startswith('As of') and value not in parts:
                        parts.append(value)
                # Identification fields take their nearest header, not a spanning title.
                headers.append(parts[-1] if c <= ids else ' / '.join(parts))
            rows, total = [], None
            for r in sorted({r for r, c in cells if r >= 4}):
                values = [cells.get((r, c)) for c in range(1, width + 1)]
                if not any(x is not None and str(x).strip() for x in values[:ids]):
                    continue
                if any(str(x).strip().casefold() in ('total', 'grand total') for x in values[:ids]):
                    total = values
                else:
                    rows.append(values)
            date = next((str(v) for (r, c), v in cells.items()
                         if r <= 3 and str(v).startswith('As of')), 'Report date unavailable')
            # Keep all columns; no hardcoded metric width.
            result[level] = dict(name=name, ids=ids, width=width, cells=cells,
                                 merges=merges, headers=headers, rows=rows,
                                 total=total, date=date, missing=missing)
            if not rows:
                raise ValueError(f'{name} has no summary rows.')
    return result


def norm(value):
    return re.sub(r'[^a-z0-9]+', '', str(value).casefold())


def indian(value, decimals=0):
    negative = '-' if value < 0 else ''
    digits = f'{abs(value):.{decimals}f}'
    whole, dot, fraction = digits.partition('.')
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while head:
            groups.insert(0, head[-2:])
            head = head[:-2]
        whole = ','.join(groups + [tail])
    return negative + whole + (dot + fraction if dot else '')


def display(value, header):
    if value is None or (isinstance(value, str) and value.strip() in ('-', '—', '')):
        return '—'
    if not isinstance(value, Real):
        return str(value)
    h = norm(header)
    if h.startswith('ftnr') and ('ftnr%' in str(header).lower() or h.endswith('ftnr')):
        return f'{value:.2%}'
    if 'grt' in h and ('grtpercent' in h or 'grt(%)' in str(header).lower() or ('%' in str(header) and h.endswith('grtpqm'))):
        return f'{value:.2%}'
    if 'dbamount' in h and 'cr' not in h:
        return '₹' + indian(value)
    if 'amountcr' in h:
        return '₹' + indian(value, 2) + ' Cr'
    if 'effectiveproductivity' in h or 'effproductivity' in h:
        return f'{value:.2%}'
    if 'productivity' in h or 'timeinminutes' in h:
        return indian(value, 2)
    return indian(value)


def render_table(info, rows, total=None):
    """Dynamic table: remove empty trailing/helper columns, preserve 3-row merged headers."""
    from math import isfinite
    headers, ids = info['headers'], info['ids']
    width = info['width']
    def numeric(x):
        try:
            if x is None or str(x).strip() in ('', '-', '—'): return None
            v = float(x)
            return v if isfinite(v) else None
        except (ValueError, TypeError): return None
    # Drop empty columns only when both the visible data AND grand total are empty.
    visible = [c for c in range(width) if c < ids or any(numeric(r[c]) is not None or
                (isinstance(r[c], str) and r[c].strip() not in ('', '-', '—')) for r in rows)
                or (total is not None and (numeric(total[c]) is not None or
                (isinstance(total[c], str) and total[c].strip() not in ('', '-', '—'))))]
    if not visible: return '<p>No data available.</p>'
    # A merged header cannot be represented as a noncontiguous span; show the
    # full header path in one sticky header row after empty columns are removed.
    def section(c):
        h = norm(headers[c])
        if c < ids: return 'identity'
        if h.startswith('ftnr'): return 'ftnr'
        if h.startswith('ujala'): return 'ujala'
        if 'freshclientgrt' in h: return 'grt'
        if 'productivity' in h or 'timeinminutes' in h: return 'productivity'
        if 'vishwas' in h or 'allocated' in h or 'pdconducted' in h: return 'vishwas'
        return 'other'
    palette = {'identity':'#243b60','vishwas':'#0369a1','grt':'#6d28d9',
               'ujala':'#c2410c','ftnr':'#047857','productivity':'#be185d','other':'#334155'}
    grt = [c for c in visible if 'freshclientgrt' in norm(headers[c]) and
           ('grt%' in headers[c].lower() or 'grt(%)' in headers[c].lower())]
    ftnr = [c for c in visible if section(c)=='ftnr' and 'ftnr%' in headers[c].lower()]
    amounts = [c for c in visible if c >= ids and 'dbamount' in norm(headers[c])]
    extrema = {}
    for c in set(grt + ftnr + amounts):
        vals = [numeric(r[c]) for r in rows]
        vals = [v for v in vals if v is not None]
        if vals: extrema[c] = (min(vals), max(vals))
    def heat(v,c):
        n=numeric(v)
        if n is None or c not in extrema:return ''
        lo,hi=extrema[c]
        t=(n-lo)/(hi-lo) if hi>lo else 1.0
        if c in ftnr:t=1-t
        if t>=.8:return 'background:#15803d;color:white;font-weight:800;'
        if t>=.6:return 'background:#bbf7d0;color:#14532d;font-weight:700;'
        if t>=.4:return 'background:#fef08a;color:#713f12;font-weight:700;'
        if t>=.2:return 'background:#fed7aa;color:#9a3412;font-weight:700;'
        return 'background:#fecaca;color:#991b1b;font-weight:700;'
    # First two identification columns remain visible while scrolling.
    frozen = visible[:min(2,ids)]
    if info['name']=='PQM Summary':
        frozen = [c for c in visible if norm(headers[c]) in ('pqmname','branch')][:2] or frozen
    def pin(c):
        if c not in frozen:return ''
        return f'position:sticky;left:{frozen.index(c)*155}px;min-width:155px;max-width:155px;z-index:3;'
    # Reconstruct merged header rows after removing blank columns. Each span counts
    # only retained columns, so filtered-out helpers never shift the heading.
    merged = {}
    for r1,c1,r2,c2 in info['merges']:
        for rr in range(r1,r2+1):
            for cc in range(c1,c2+1):
                merged[rr,cc]=(r1,c1,r2,c2)
    out=['<div class="report-scroll"><table class="report"><thead>']
    occupied=set()
    for rr in range(1,4):
        out.append('<tr>')
        for c in visible:
            if (rr,c) in occupied:continue
            orig=merged.get((rr,c+1))
            if orig:
                r1,c1,r2,c2=orig
                cols=[v for v in visible if c1<=v+1<=c2]
                if not cols:continue
                if rr!=r1 or c!=cols[0]:continue
                span=len(cols);rowspan=r2-r1+1
            else:
                span=rowspan=1
                cols=[c]
            for rrr in range(rr,rr+rowspan):
                for cc in cols:occupied.add((rrr,cc))
            value=info['cells'].get((rr,c+1),'')
            label=escape(str(value if value is not None else '')).replace('\n','<br>')
            group=section(c)
            style=f'background:{palette[group]};min-width:120px;max-width:190px;white-space:normal;overflow-wrap:break-word;'
            if c in frozen:style+=pin(c)+'z-index:8;'
            out.append(f'<th rowspan="{rowspan}" colspan="{span}" style="{style}">{label}</th>')
        out.append('</tr>')
    out.append('</thead><tbody>')
    for ri,row in enumerate(rows):
        out.append('<tr>')
        for c in visible:
            value = escape(str(row[c]) if row[c] is not None else '—') if c<ids else escape(display(row[c],headers[c]))
            bg = '#ffffff' if ri%2==0 else '#f0f5fc'
            style = f'background:{bg};'+pin(c) if c in frozen else heat(row[c],c)
            out.append(f'<td class="{"label" if c<ids else "number"}" style="{style}">{value}</td>')
        out.append('</tr>')
    if total is not None:
        out.append('<tr class="total">')
        for c in visible:
            value=escape(str(total[c]) if total[c] is not None else '—') if c<ids else escape(display(total[c],headers[c]))
            out.append(f'<td style="background:#172f53;color:white;font-weight:800;{pin(c)}">{value}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div>')
    return ''.join(out)


def normalize_filter_header(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())

def main():
    import streamlit as st
    st.set_page_config(page_title='PQM Dashboard', layout='wide')
    st.markdown('''<style>
    .stApp{background:#f4f7fb}.block-container{padding-top:2rem}
    .hero{background:linear-gradient(100deg,#102b50,#0b7496);color:white;padding:24px 30px;border-radius:16px;margin-bottom:20px}
    .kpi-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:8px;margin:10px 0 14px}.kpi-tile{background:white;border:1px solid #dce5ef;border-top:3px solid #167b9d;border-radius:8px;padding:8px 10px;min-width:0}.kpi-group-label{font-size:10px;line-height:14px;color:#167b9d;font-weight:700}.kpi-label{font-size:11px;line-height:14px;color:#536579;min-height:28px;overflow-wrap:anywhere}.kpi-number{font-size:clamp(15px,1.45vw,21px);line-height:26px;font-weight:700;color:#123956;overflow-wrap:anywhere}
    @media(max-width:900px){.kpi-grid{grid-template-columns:repeat(3,minmax(0,1fr))}}
    @media(max-width:480px){.kpi-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}

    .compact-report-heading{display:flex;align-items:center;gap:22px;flex-wrap:wrap;margin:7px 0 10px;color:#64748b;font-size:13px}.compact-report-heading strong{font-size:24px;color:#172b4d} .hero h1{margin:0;font-size:30px}.hero p{margin:6px 0 0;color:#deedf4}
    .report-scroll{overflow:auto;max-height:670px;border:1px solid #d7e0eb;border-radius:10px;background:white}
    .gauge-card{background:white;border:1px solid #dbeafe;border-radius:18px;padding:10px 14px;box-shadow:0 7px 22px #22386010} .progress-panel{background:white;border:1px solid #dbeafe;border-radius:18px;padding:26px;margin-top:14px;box-shadow:0 7px 22px #22386010}
    .progress-line{display:flex;justify-content:space-between;gap:12px;font-size:14px;margin:16px 0 8px;color:#334155}
    .progress-track{height:18px;border-radius:30px;background:#e2e8f0;overflow:hidden}
    .progress-fill{height:100%;border-radius:30px}
    .conversion-footer{margin-top:24px;padding-top:16px;border-top:1px solid #e2e8f0;color:#475569}
    .conversion-footer strong{color:#059669;font-size:22px;margin-left:8px}
    .report{border-collapse:separate;border-spacing:0;font:12px Arial;min-width:100%}
    .report th{background:#133d65;color:white;min-width:120px;padding:11px 9px;border-right:1px solid #4b6984;border-bottom:1px solid #4b6984;text-align:center}
    .report thead{position:sticky;top:0;z-index:2}.report td{padding:9px 12px;border-right:1px solid #e0e6ef;border-bottom:1px solid #e0e6ef;white-space:nowrap;text-align:right}
    .report td.label{text-align:left}.report tbody tr:nth-child(even){background:#f0f5fa}
    .report tbody tr:hover{outline:2px solid #60a5fa;outline-offset:-2px}
    .report .frozen{position:sticky;left:0;text-align:left;white-space:normal;overflow-wrap:anywhere;box-shadow:2px 0 3px #c4d3e3}
    .report td.frozen{z-index:1;background:#fff}
    .report tbody tr:nth-child(even) td.frozen{background:#f0f5fa}
    .report tbody tr:hover td.frozen{background:#e2f1fa}
    .report th.frozen{z-index:4;background:#133d65}
    .report tr.total{background:#172f53;color:white;font-weight:bold}
    </style>''', unsafe_allow_html=True)
    st.markdown('<div class="hero"><h1>Welcome to the PQM Dashboard Analyser</h1><p>Explore your performance in visual and detailed table views. Upload your latest report.</p></div>', unsafe_allow_html=True)

    upload = st.file_uploader(
        "Upload your Combined PQM Report (.xlsx)",
        type=["xlsx"],
        help="Choose the report you want to analyse. Your upload is used for this session only."
    )
    if upload is None:
        st.info("Upload a report to get started. No data or analysis is displayed until a file is uploaded.")
        st.stop()

    raw = upload.getvalue()
    try:
        # Deliberately do not cache report contents, to avoid retaining confidential
        # report data in Streamlit's shared data cache.
        reports = load_report(raw)
    except Exception as e:
        st.error(f"Could not open this report: {e}")
        st.info("Please upload a valid Combined PQM Report (.xlsx) with the required summary sheets.")
        st.stop()

    st.sidebar.title("Report controls")
    level = st.sidebar.selectbox("Summary", list(SHEETS))
    info = reports[level]
    st.markdown(f"""<div class='compact-report-heading'><strong>{escape(info['name'])}</strong><span>📅 {escape(info['date'])}</span><span>📄 {escape(upload.name)}</span></div>""", unsafe_allow_html=True)

    if info['missing']:
        st.error('Some formula results are missing. Open the workbook in Excel, recalculate, save, and upload again.')
        st.stop()
    df = pd.DataFrame(info['rows'], columns=info['headers'])
    filtered = df.copy()
    entity_filters = {}
    branch_detail = False
    for c in range(info['ids']):
        # Counts are metrics, not entity filters.
        if normalize_filter_header(info['headers'][c]).startswith('no'):
            continue
        series = filtered.iloc[:, c].fillna('').astype(str)
        choices = sorted(x for x in series.unique() if x)
        selected = st.sidebar.multiselect(info['headers'][c], choices, key=f'{level}-{c}')
        if selected:
            entity_filters[info['headers'][c]] = selected
            filtered = filtered[series.isin(selected)]
    if level in ('Zone', 'Region(AVP)', 'DM'):
        pq_info = reports['PQM']
        branch_rows = pd.DataFrame(pq_info['rows'], columns=pq_info['headers'])
        for field, chosen in entity_filters.items():
            field = 'RM Name' if field == 'RM' else field
            if field in branch_rows.columns:
                branch_rows = branch_rows[branch_rows[field].fillna('').astype(str).isin(chosen)]
        branches = sorted(x for x in branch_rows['Branch'].dropna().astype(str).unique() if x.strip() and x != '-')
        selected_branches = st.sidebar.multiselect('PQ Branch', branches, key=f'{level}-branch', disabled=(level == 'DM' and not entity_filters.get('DM Name')))
        if selected_branches or (level == 'Region(AVP)' and entity_filters.get('Region AVP')) or (level == 'DM' and entity_filters.get('DM Name')):
            branch_rows = branch_rows[branch_rows['Branch'].fillna('').astype(str).str.strip().ne('') & branch_rows['Branch'].fillna('').astype(str).str.strip().ne('-')]
            if pq_info['missing']:
                st.error('PQM formula results are missing. Recalculate and save the Excel report first.')
                st.stop()
            info = pq_info
            filtered = (branch_rows[branch_rows['Branch'].fillna('').astype(str).isin(selected_branches)] if selected_branches else branch_rows).copy()
            branch_detail = True
            st.caption('PQ Branch Performance · Source: PQM Summary')
    all_rows = not branch_detail and len(filtered) == len(df)
    def sum_col(offset):
        if filtered.empty:
            return None
        if all_rows and info['total'] is not None:
            total_value = info['total'][info['ids']+offset]
            if isinstance(total_value, Real):
                return total_value
        raw_values = filtered.iloc[:, info['ids']+offset]
        # Excel uses '-' for non-applicable counts (e.g. Other Branches GRT).
        # Exclude that explicit marker, but keep real missing/error values unknown.
        applicable = raw_values[~raw_values.map(lambda x: isinstance(x, str) and x.strip() == '-')]
        s = pd.to_numeric(applicable, errors='coerce')
        return s.sum() if len(s) and s.notna().all() else None
    # Match full metric paths, never fixed Excel column offsets.
    # A renamed/ambiguous KPI is shown as unavailable rather than a wrong number.
    def normalize(x):
        return re.sub(r'[^a-z0-9]+', '', str(x).casefold())
    def find_metric(path):
        target = normalize(path)
        hits = [i for i,h in enumerate(info['headers']) if normalize(h).endswith(target)]
        return hits[0] if len(hits) == 1 else None

    SPECS = [
        ('Vishwas', '#Total Allocated', 'count', '#Total Allocated'),
        ('Vishwas', 'PD Conducted', 'count', '# Total PD Conducted'),
        ('Vishwas', 'DB', 'count', 'Vishwas Pipeline / # DB'),
        ('Vishwas', 'DB Amount (Cr)', 'crore', '#DB Amount(Cr)'),
        ('Vishwas', 'Declined', 'count', 'Vishwas Pipeline / Declined'),
        ('GRT', 'PQM GRT', 'count', 'Fresh Client GRT / PQM GRT'),
        ('GRT', 'GRT % PQM', 'grt_rate', 'Fresh Client GRT / GRT(%) PQM'),
        ('Ujala', 'CB to DB', 'count', 'Ujala / CB to DB'),
        ('Ujala', 'DB Amount', 'rupees', 'Ujala / DB Amount'),
        ('FTNR', 'FTNR %', 'ftnr_rate', 'FTNR %'),
        ('Productivity', 'Effective Productivity', 'weighted_rate', 'Time in Minutes / Effective Productivity'),
    ]
    def locate(key):
        if key == 'Vishwas Pipeline / # DB':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('vishwaspipeline') and normalize(h).endswith('db') and not normalize(h).endswith('amountcr')]
            return hits[0] if len(hits)==1 else None
        if key == 'Vishwas Pipeline / Declined':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('vishwaspipeline') and normalize(h).endswith('declined')]
            return hits[0] if len(hits)==1 else None
        if key == 'Time in Minutes / Effective Productivity':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('timeinminutes') and (normalize(h).endswith('effproductivity') or normalize(h).endswith('effectiveproductivity'))]
            return hits[0] if len(hits)==1 else None
        if key == 'FTNR %':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('ftnr') and normalize(h).endswith('ftnr') and 'totalgrt' not in normalize(h)]
            return hits[0] if len(hits)==1 else None
        return find_metric(key)
    def aggregate(idx):
        if idx is None or filtered.empty: return None
        if all_rows and info['total'] is not None:
            v=info['total'][idx]
            if isinstance(v,Real):return float(v)
        series=filtered.iloc[:,idx]
        series=series[~series.map(lambda x:isinstance(x,str) and x.strip()=='-')]
        numeric=pd.to_numeric(series,errors='coerce')
        return float(numeric.sum()) if len(numeric) and numeric.notna().all() else None
    def metric(spec):
        group,label,kind,key=spec
        idx=locate(key)
        if idx is None:return None
        if kind=='grt_rate':
            a,b=aggregate(locate('Fresh Client GRT / PQM GRT')),aggregate(locate('Fresh Client GRT / Total'))
            return a/b if a is not None and b else None
        if kind=='ftnr_rate':
            a,b=aggregate(locate('FTNR / PQM(Regular + Ujala +Physical PD)')),aggregate(locate('FTNR / PQM(Regular + Ujala +Physical PD) / Total GRT'))
            return a/b if a is not None and b else None
        if kind=='weighted_rate':
            if all_rows and info['total'] is not None and isinstance(info['total'][idx],Real):return float(info['total'][idx])
            vals=pd.to_numeric(filtered.iloc[:,idx],errors='coerce')
            if level=='PQM' or branch_detail:
                weights=pd.Series(1.0,index=filtered.index)
                emp_col=next((i for i,h in enumerate(info['headers']) if normalize(h)=='pqmempid'),None)
                if emp_col is not None:weights=filtered.iloc[:,emp_col].fillna('').astype(str).str.strip().map(lambda x:float(bool(x and x!='-')))
            else:
                count_col=next((i for i,h in enumerate(info['headers']) if normalize(h) in ('noofpqm','noofpqms')),None)
                if count_col is None:return None
                weights=pd.to_numeric(filtered.iloc[:,count_col],errors='coerce')
            valid=(weights>0)&weights.notna()&vals.notna()
            return float((vals[valid]*weights[valid]).sum()/weights[valid].sum()) if valid.any() else None
        return aggregate(idx)
    def fmt(v,kind):
        if v is None or pd.isna(v):return '—'
        if kind in ('grt_rate','ftnr_rate'):return f'{v:.2%}'
        if kind=='weighted_rate':return f'{v:.2%}'
        if kind=='crore':return f'₹{v:,.2f} Cr'
        if kind=='rupees':return f'₹{v/1e7:,.2f} Cr'
        return f'{v:,.0f}'
    view=st.radio('Display mode',['📊 Visual Dashboard','📋 Table View'],horizontal=True)
    palette={'Vishwas':'#0284c7','GRT':'#7c3aed','Ujala':'#ea580c','FTNR':'#059669','Productivity':'#db2777'}
    st.markdown("""<style>.stApp{background:linear-gradient(125deg,#f0f7ff,#f9f5ff)}
    [data-testid="stMetric"]{background:white;border:1px solid #dbeafe;border-radius:14px;padding:16px;box-shadow:0 5px 16px #c8d7ec50}
    [data-testid="stMetricLabel"]{color:#475569}</style>""",unsafe_allow_html=True)
    if view=='📊 Visual Dashboard':
        import plotly.express as px
        import plotly.graph_objects as go
        st.markdown("""<style>
        .block-container{max-width:1550px;padding-top:1.4rem}
        .visual-heading{font-size:19px;font-weight:800;color:#172b4d;margin:24px 0 12px}
        .visual-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:12px 0 20px}
        .visual-card{background:white;border:1px solid #e5eaf4;border-radius:18px;padding:19px 18px 17px;
          box-shadow:0 7px 22px rgba(34,56,96,.065);border-top:5px solid var(--accent);min-width:0}
        .visual-icon{font-size:22px;background:var(--soft);border-radius:11px;width:42px;height:42px;
          display:flex;align-items:center;justify-content:center;margin-bottom:13px}
        .visual-label{font-size:13px;font-weight:600;color:#64748b;min-height:19px}
        .visual-value{font-size:clamp(20px,2.3vw,32px);font-weight:800;color:#15243d;letter-spacing:-.8px;margin-top:7px;overflow-wrap:anywhere}
        .visual-foot{font-size:11px;color:#94a3b8;margin-top:8px}
        @media(max-width:1000px){.visual-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}
        @media(max-width:530px){.visual-grid{grid-template-columns:1fr}}
        .paired-sections{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:22px;margin-bottom:18px;align-items:start}
        .paired-sections:has(> .paired-section:only-child){grid-template-columns:1fr}
        .paired-cards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
        .paired-section .visual-heading{margin:12px 0 10px}
        .paired-section .visual-card{padding:14px 15px;min-height:145px}
        .paired-section .visual-icon{width:32px;height:32px;font-size:18px;margin-bottom:8px}
        .paired-section .visual-value{font-size:clamp(18px,2vw,27px)}
        @media(max-width:900px){.paired-sections{grid-template-columns:1fr}}
        @media(max-width:480px){.paired-cards{grid-template-columns:1fr 1fr;gap:8px}.paired-section .visual-card{padding:10px}}
        </style>""",unsafe_allow_html=True)
        st.markdown('### Performance overview')
        st.caption('Live view of the selected Excel summary · Filters apply to cards and charts')
        icons={'Vishwas':'💼','GRT':'🎯','Ujala':'✨','FTNR':'🛡️','Productivity':'⏱️'}
        soft={'Vishwas':'#e0f2fe','GRT':'#ede9fe','Ujala':'#ffedd5','FTNR':'#d1fae5','Productivity':'#fce7f3'}
        grouped={g:[sp for sp in SPECS if sp[0]==g] for g in palette}
        # Compact paired KPI sections: two categories per row.
        for pair in [('Vishwas',), ('GRT', 'FTNR'), ('Ujala', 'Productivity')]:
            sections = ['<div class="paired-sections">']
            for group in pair:
                sections.append(f'<section class="paired-section"><div class="visual-heading"><span style="color:{palette[group]}">{icons[group]}</span> {escape(group)} Performance</div><div class="paired-cards">')
                for spec in grouped[group]:
                    _,label,kind,key=spec
                    val=fmt(metric(spec),kind)
                    sections.append(f'<div class="visual-card" style="--accent:{palette[group]};--soft:{soft[group]}"><div class="visual-icon">{icons[group]}</div><div class="visual-label">{escape(label)}</div><div class="visual-value">{escape(val)}</div></div>')
                sections.append('</div></section>')
            sections.append('</div>')
            st.markdown(''.join(sections),unsafe_allow_html=True)
        st.markdown('### Performance analytics')
        def chart_frame(key):
            """At DM/RM selection drill down to PQM Summary branches, not repeated DM labels."""
            source = filtered
            source_info = info
            selected_dm = entity_filters.get('DM Name', [])
            selected_rm = entity_filters.get('RM', []) or entity_filters.get('RM Name', [])
            if level in ('Zone','Region(AVP)','DM') and (selected_dm or selected_rm or branch_detail):
                source_info = reports['PQM']
                source = pd.DataFrame(source_info['rows'],columns=range(source_info['width']))
                for field, chosen in entity_filters.items():
                    alias = {'RM':'RM Name','Region AVP':'Region AVP'}.get(field,field)
                    candidates=[i for i,h in enumerate(source_info['headers']) if norm(h)==norm(alias)]
                    if candidates and chosen:
                        source=source[source.iloc[:,candidates[0]].fillna('').astype(str).isin(chosen)]
                if selected_branches:
                    branch_idx=next((i for i,h in enumerate(source_info['headers']) if norm(h)=='branch'),None)
                    if branch_idx is not None:source=source[source.iloc[:,branch_idx].astype(str).isin(selected_branches)]
            metric_hits=[i for i,h in enumerate(source_info['headers']) if norm(h).endswith(norm(key))]
            if len(metric_hits)!=1:
                if key=='Vishwas Pipeline / # DB':
                    metric_hits=[i for i,h in enumerate(source_info['headers']) if norm(h).startswith('vishwaspipeline') and norm(h).endswith('db')]
            if len(metric_hits)!=1:return None
            # Use branch labels when drilling into DM, otherwise appropriate entity.
            preferred = (['Branch'] if source_info['name']=='PQM Summary' else
                         ['DM Name','Region AVP Name','Zonal Name','Zone'])
            name_idx=next((i for name in preferred for i,h in enumerate(source_info['headers'][:source_info['ids']]) if norm(h)==norm(name)),0)
            if isinstance(source.columns[0],int):
                names=source.iloc[:,name_idx]
                values=source.iloc[:,metric_hits[0]]
            else:
                names=source.iloc[:,name_idx]
                values=source.iloc[:,metric_hits[0]]
            data=pd.DataFrame({'Name':names.astype(str).str.strip().to_numpy(),
                               'Value':pd.to_numeric(values,errors='coerce').to_numpy()})
            data=data[(data.Name!='') & (data.Name!='-')].dropna(subset=['Value'])
            if data.empty:return None
            return data.groupby('Name',as_index=False)['Value'].sum().sort_values('Value',ascending=False).head(10)

        def ranked_chart(column,key,heading,colors):
            with column:
                st.markdown('**'+heading+'**')
                data=chart_frame(key)
                if data is None or data.empty:
                    st.info('Chart data unavailable for this selection.')
                    return
                data=data.sort_values('Value',ascending=True)
                fig=px.bar(data,x='Value',y='Name',orientation='h',color='Value',
                           color_continuous_scale=colors,text='Value',
                           hover_data={'Name':True,'Value':':,.0f'})
                vmax=max(float(data.Value.max()),1)
                fig.update_xaxes(range=[0,vmax*1.28],showgrid=True,gridcolor='#e7eaf2',title=None)
                fig.update_yaxes(title=None,automargin=True)
                fig.update_traces(texttemplate='%{x:,.0f}',textposition='outside',cliponaxis=False)
                fig.update_layout(height=max(360,55*len(data)+100),coloraxis_showscale=False,
                                  margin=dict(l=15,r=80,t=18,b=20),
                                  paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                                  font=dict(size=12,color='#334155'),showlegend=False)
                st.plotly_chart(fig,use_container_width=True,config={'displayModeBar':False})
        left,right=st.columns(2,gap='large')
        ranked_chart(left,'Vishwas Pipeline / # DB','Top 10 · Disbursement Count',['#bae6fd','#0284c7','#075985'])
        ranked_chart(right,'Fresh Client GRT / PQM GRT','Top 10 · PQM GRT',['#ddd6fe','#8b5cf6','#5b21b6'])
        left,right=st.columns(2,gap='large')
        with left:
            st.markdown('**PQM GRT Performance · Speedometer**')
            grt_rate=metric(next(sp for sp in SPECS if sp[2]=='grt_rate'))
            if grt_rate is None:
                st.info('PQM GRT % unavailable.')
            else:
                pct=grt_rate*100
                from math import pi, cos, sin
                # SVG semi-circle: compact, responsive and accurately scaled 0–100.
                cx,cy,r=200,174,122
                arc_length=pi*r
                progress=max(0,min(100,pct))
                svg=f"""<div class='gauge-card'><svg viewBox='0 0 400 228' width='100%' style='max-height:230px'>
                <defs><linearGradient id='pqmArc' x1='0%' y1='0%' x2='100%' y2='0%'>
                <stop offset='0%' stop-color='#f59e0b'/><stop offset='55%' stop-color='#8b5cf6'/><stop offset='100%' stop-color='#059669'/></linearGradient></defs>
                <path d='M 78 174 A 122 122 0 0 1 322 174' stroke='#e2e8f0' stroke-width='24' stroke-linecap='round' fill='none'/>
                <path d='M 78 174 A 122 122 0 0 1 322 174' stroke='url(#pqmArc)' stroke-width='24' stroke-linecap='round' fill='none'
                stroke-dasharray='{arc_length*progress/100:.3f} {arc_length:.3f}'/>
                <text x='200' y='151' text-anchor='middle' font-size='42' font-weight='800' fill='#172b4d'>{pct:.2f}%</text>
                <text x='200' y='177' text-anchor='middle' font-size='14' fill='#64748b'>PQM GRT</text>
                <text x='70' y='209' font-size='14' fill='#64748b'>0%</text><text x='291' y='209' font-size='14' fill='#64748b'>100%</text>
                </svg></div>"""
                st.markdown(svg,unsafe_allow_html=True)
                if pct>100:st.caption('Actual value exceeds the 100% gauge maximum.')
        with right:
            st.markdown('**Disbursement Performance · Progress Comparison**')
            allocated=aggregate(locate('#Total Allocated'))
            conducted=aggregate(locate('# Total PD Conducted'))
            disbursed=aggregate(locate('Vishwas Pipeline / # DB'))
            if all(v is not None for v in (allocated,conducted,disbursed)):
                values=[allocated,conducted,disbursed]
                maximum=max(max(values),1)
                labels=['Allocated','PD Conducted','Disbursed']
                colors=['#38bdf8','#8b5cf6','#10b981']
                html=['<div class="progress-panel">']
                for label,value,color in zip(labels,values,colors):
                    html.append(f'<div class="progress-line"><span>{escape(label)}</span><strong>{indian(value)}</strong></div>')
                    html.append(f'<div class="progress-track"><div class="progress-fill" style="width:{max(0,min(100,value/maximum*100)):.2f}%;background:{color}"></div></div>')
                rate=100*disbursed/conducted if conducted else None
                html.append('<div class="conversion-footer">PD to DB Conversion: <strong>'+ (f'{rate:.1f}%' if rate is not None else '—') +'</strong></div></div>')
                st.markdown(''.join(html),unsafe_allow_html=True)
            else:st.info('Conversion data unavailable.')
    else:
        st.caption('🟢 Green = stronger performance · 🟡 Yellow = middle · 🔴 Red = needs attention. GRT: higher is better; FTNR: lower is better. Colours are relative to filtered rows.')
        cards=['<div class="kpi-grid">']
        for group,label,kind,key in SPECS:
            cards.append(f'<div class="kpi-tile"><div class="kpi-group-label">{escape(group)}</div><div class="kpi-label">{escape(label)}</div><div class="kpi-number">{escape(fmt(metric((group,label,kind,key)),kind))}</div></div>')
        cards.append('</div>')
        st.markdown(''.join(cards),unsafe_allow_html=True)
    missing_kpis=[spec[1] for spec in SPECS if locate(spec[3]) is None]
    if missing_kpis:st.warning('These KPIs could not be matched uniquely to Excel headers: '+', '.join(missing_kpis)+'. They are shown as —, not guessed.')
    st.caption(f'{len(filtered):,} branch detail rows · Scroll horizontally to see all metrics.' if branch_detail else f'{len(filtered):,} of {len(df):,} rows · Scroll horizontally to see all metrics.')
    if filtered.empty:
        st.info('No matching rows. Clear the filters.')
    elif view == '📋 Table View':
        st.markdown(render_table(info, filtered.values.tolist(), info['total'] if all_rows else None), unsafe_allow_html=True)
    st.caption('Figures come directly from the selected Excel summary. The Excel total is shown only for the full view. Filtered percentages are not averaged.')



if __name__ == '__main__':
    main()
