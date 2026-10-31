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
            width = max((colnum(c.attrib['r']) for c in root.findall('s:sheetData/s:row/s:c', NS)), default=0)
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


def display(value, header):
    if value is None: return '—'
    if not isinstance(value, Real): return str(value)
    h = str(header).lower()
    if 'ftnr %' in h or 'grt(%)' in h or 'eff productivity' in h or 'effective productivity' in h:
        return f'{value:.2%}'
    if 'productivity' in h or 'time in minutes' in h or 'amount(cr)' in h:
        return f'{value:,.2f}'
    return f'{value:,.0f}'


def render_table(info, rows, total=None):
    frozen = {'Zone Summary': [1], 'Region(AVP) Summary': [1, 2],
              'DM Summary': [1, 2], 'PQM Summary': [8]}[info['name']]
    def pin(column):
        if column not in frozen:
            return ''
        left = frozen.index(column) * 160
        return f' class="frozen" style="left:{left}px;min-width:160px;width:160px;max-width:160px;box-sizing:border-box"'
    covered, origins = set(), {}
    for r1, c1, r2, c2 in info['merges']:
        origins[r1, c1] = (r2-r1+1, c2-c1+1)
        covered.update((r, c) for r in range(r1, r2+1) for c in range(c1, c2+1)
                       if (r, c) != (r1, c1))
    out = ['<div class="report-scroll"><table class="report"><thead>']
    for r in range(1, 4):
        out.append('<tr>')
        for c in range(1, info['width']+1):
            if (r, c) in covered:
                continue
            rs, cs = origins.get((r, c), (1, 1))
            text = escape(str(info['cells'].get((r, c), ''))).replace('\n', '<br>')
            out.append(f'<th{pin(c)} rowspan="{rs}" colspan="{cs}">{text}</th>')
        out.append('</tr>')
    out.append('</thead><tbody>')
    for values in rows:
        out.append('<tr>')
        for c, v in enumerate(values):
            label = c < info['ids']
            attrs = pin(c+1) if c+1 in frozen else f' class="{"label" if label else "number"}"'
            out.append(f'<td{attrs}>{escape(str(v or "") if label else display(v,info["headers"][c]))}</td>')
        out.append('</tr>')
    if total:
        out.append('<tr class="total">')
        for c, v in enumerate(total):
            out.append(f'<td{pin(c+1)}>{escape(str(v or "") if c < info["ids"] else display(v,info["headers"][c]))}</td>')
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

    .hero h1{margin:0;font-size:30px}.hero p{margin:6px 0 0;color:#deedf4}
    .report-scroll{overflow:auto;max-height:670px;border:1px solid #d7e0eb;border-radius:10px;background:white}
    .report{border-collapse:separate;border-spacing:0;font:12px Arial;min-width:100%}
    .report th{background:#133d65;color:white;min-width:95px;padding:11px 9px;border-right:1px solid #4b6984;border-bottom:1px solid #4b6984;text-align:center}
    .report thead{position:sticky;top:0;z-index:2}.report td{padding:9px 12px;border-right:1px solid #e0e6ef;border-bottom:1px solid #e0e6ef;white-space:nowrap;text-align:right}
    .report td.label{text-align:left}.report tbody tr:nth-child(even){background:#f0f5fa}
    .report tbody tr:hover{background:#e2f1fa}
    .report .frozen{position:sticky;left:0;text-align:left;white-space:normal;overflow-wrap:anywhere;box-shadow:2px 0 3px #c4d3e3}
    .report td.frozen{z-index:1;background:#fff}
    .report tbody tr:nth-child(even) td.frozen{background:#f0f5fa}
    .report tbody tr:hover td.frozen{background:#e2f1fa}
    .report th.frozen{z-index:4;background:#133d65}
    .report tr.total td.frozen{background:#d9e8f5}.report tr.total{background:#d9e8f5;font-weight:bold}
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
    st.subheader(info['name'])
    st.caption(info['date'])
    st.caption(f"Analysis of uploaded file: {escape(upload.name)}")

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
            st.caption('PQ branch performance for the selected filters · Source: PQM Summary.')
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
        ('Productivity', 'Effective Productivity', 'weighted_rate', 'Time in Minutes / Eff Productivity'),
    ]
    def locate(key):
        if key == 'Vishwas Pipeline / # DB':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('vishwaspipeline') and normalize(h).endswith('db') and not normalize(h).endswith('amountcr')]
            return hits[0] if len(hits)==1 else None
        if key == 'Vishwas Pipeline / Declined':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('vishwaspipeline') and normalize(h).endswith('declined')]
            return hits[0] if len(hits)==1 else None
        if key == 'Time in Minutes / Eff Productivity':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('timeinminutes') and (normalize(h).endswith('effproductivity') or normalize(h).endswith('effectiveproductivity'))]
            return hits[0] if len(hits)==1 else None
        if key == 'FTNR %':
            hits=[i for i,h in enumerate(info['headers']) if normalize(h).startswith('ftnr') and normalize(h).endswith('ftnr')]
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
        if kind in ('grt_rate','ftnr_rate','weighted_rate'):return f'{v:.2%}'
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
        for group in palette:
            st.markdown(f'<h3 style="color:{palette[group]}">{group}</h3>',unsafe_allow_html=True)
            group_specs=[x for x in SPECS if x[0]==group]
            cols=st.columns(len(group_specs))
            for col,spec in zip(cols,group_specs):col.metric(spec[1],fmt(metric(spec),spec[2]))
        if not filtered.empty:
            st.subheader('Performance comparison')
            chart_choices=[('DB','Vishwas Pipeline / # DB'),('PQM GRT','Fresh Client GRT / PQM GRT'),('FTNR %','FTNR %')]
            choice=st.selectbox('Compare by',[x[0] for x in chart_choices])
            idx=locate(dict(chart_choices)[choice])
            if idx is not None:
                entity_col=next((i for i,h in enumerate(info['headers']) if normalize(h) in ('dmname','pqmname','zonalname','regionavpname','branch')),0)
                plot=pd.DataFrame({'Entity':filtered.iloc[:,entity_col].fillna('Unknown').astype(str).to_numpy(),'Value':pd.to_numeric(filtered.iloc[:,idx],errors='coerce').to_numpy()})
                plot=plot.dropna(subset=['Value'])
                if choice=='FTNR %':plot['Value']*=100
                if not plot.empty:
                    top=plot.sort_values('Value',ascending=(choice=='FTNR %')).head(10)
                    fig=px.bar(top,x='Value',y='Entity',orientation='h',color='Value',color_continuous_scale='Blues')
                    fig.update_layout(height=430,yaxis={'categoryorder':'total ascending'},margin=dict(l=10,r=10,t=20,b=20))
                    st.plotly_chart(fig,use_container_width=True)
            else:st.info('Chart metric is not available in this sheet.')
    else:
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
