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
SHEETS = {'Zone': ('Zone Summary', 3, 39),
          'Region(AVP)': ('Region(AVP) Summary', 4, 40),
          'DM': ('DM Summary', 4, 38),
          'PQM': ('PQM Summary', 8, 42)}
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
        for level, (name, ids, width) in SHEETS.items():
            if name not in sheets:
                raise ValueError(f'Missing sheet: {name}. Upload the Combined PQM Report .xlsx.')
            root = ET.fromstring(z.read(sheets[name]))
            cells, missing = {}, 0
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
            result[level] = dict(name=name, ids=ids, width=width, cells=cells,
                                 merges=merges, headers=headers, rows=rows,
                                 total=total, date=date, missing=missing)
            if not rows:
                raise ValueError(f'{name} has no summary rows.')
    return result


def display(value, offset):
    if value is None:
        return '—'
    if not isinstance(value, Real):
        return str(value)
    if offset in (22, 30, 36, 19, 27, 33):
        return f'{value:.2%}'
    if offset in (4, 7, 11, 18, 23, 28, 29, 30, 31, 32):
        return f'{value:,.2f}'
    if offset == 26:
        return f'₹ {value:,.0f}'
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
            out.append(f'<td{attrs}>{escape(str(v or "") if label else display(v,c-info["ids"]))}</td>')
        out.append('</tr>')
    if total:
        out.append('<tr class="total">')
        for c, v in enumerate(total):
            out.append(f'<td{pin(c+1)}>{escape(str(v or "") if c < info["ids"] else display(v,c-info["ids"]))}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div>')
    return ''.join(out)


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
        if info['headers'][c].lower().startswith('no'):
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
    # Counts and amounts are additive; rates are ratios of matching totals.
    # All offsets are relative to each sheet's identification columns.
    # The new October report adds two metric columns in Zone/Region.
    # Metric offsets are consistent across the four summary sheets.
    groups = [
        ('Vishwas', [('Total Allocated', 0), ('PD Conducted', 1), ('DB', 13), ('DB Amount (Cr)', 14), ('Declined', 15)]),
        ('GRT', [('PQM GRT', 19), ('GRT % PQM', 22)]),
        ('Ujala', [('CB to DB', 23), ('DB Amount', 27)]),
        ('FTNR', [('FTNR %', 30)]),
        ('Productivity', [('Effective Productivity', 36)])]
    shift = 3 if level in ('DM', 'PQM') else 0
    groups = [(g, [(label, off-shift) for label, off in specs]) for g, specs in groups]
    # In the DM and PQM sheets the identifiers occupy 4 and 8 columns,
    # and the metrics above still use the same relative offsets.
    def metric(offset):
        grt_off, grt_total_off, ftnr_off, ftnr_total_off, prod_off = (22-shift, 20-shift, 30-shift, 29-shift, 36-shift)
        if offset == grt_off:
            a, b = sum_col(19-shift), sum_col(grt_total_off)
            return a / b if a is not None and b else None
        if offset == ftnr_off:
            a, b = sum_col(28-shift), sum_col(ftnr_total_off)
            return a / b if a is not None and b else None
        if offset == prod_off:
            if all_rows and info['total'] is not None:
                v = info['total'][info['ids'] + offset]
                return v if isinstance(v, Real) else None
            vals = pd.to_numeric(filtered.iloc[:, info['ids'] + offset], errors='coerce')
            if len(vals) == 1:
                return vals.iloc[0]
            if info['name'] == 'PQM Summary':
                weights = filtered.iloc[:, 6].fillna('').astype(str).str.strip().map(lambda x: int(bool(x and x != '-')))
            else:
                weights = pd.to_numeric(filtered.iloc[:, info['ids']-2], errors='coerce').fillna(0)
            valid = (weights > 0) & vals.notna()
            return (vals[valid] * weights[valid]).sum() / weights[valid].sum() if valid.any() else None
        return sum_col(offset)
    def fmt(v, off):
        if v is None or pd.isna(v): return '—'
        if off in (22-shift,30-shift,36-shift): return f'{v:.1%}'
        if off == 14-shift: return f'₹{v:,.2f} Cr'
        if off == 27-shift: return f'₹{v/1e7:,.2f} Cr'
        return f'{v:,.0f}'
    view = st.radio('Display mode', ['📊 Visual Dashboard', '📋 Table View'], horizontal=True)
    colors = ['#0ea5e9','#8b5cf6','#f59e0b','#10b981','#ec4899']
    if view == '📊 Visual Dashboard':
        import plotly.express as px
        st.markdown('<style>.stApp{background:linear-gradient(135deg,#f0f7ff,#faf5ff)}[data-testid="stMetric"]{background:white;border:1px solid #dbeafe;border-radius:14px;padding:16px;box-shadow:0 4px 15px #dbeafe70}</style>', unsafe_allow_html=True)
        for (group, specs), color in zip(groups, colors):
            st.markdown(f'<h3 style="color:{color};margin-bottom:8px">{group}</h3>', unsafe_allow_html=True)
            cols = st.columns(len(specs))
            for col, (label, off) in zip(cols, specs):
                col.metric(label, fmt(metric(off), off))
        if not filtered.empty:
            labels = filtered.iloc[:, 1 if info['ids'] > 1 else 0].fillna('Unknown').astype(str)
            chart_df = pd.DataFrame({'Entity': labels, 'DB': pd.to_numeric(filtered.iloc[:, info['ids']+13-shift], errors='coerce').fillna(0), 'PQM GRT': pd.to_numeric(filtered.iloc[:, info['ids']+19-shift], errors='coerce').fillna(0), 'FTNR %': pd.to_numeric(filtered.iloc[:, info['ids']+30-shift], errors='coerce')*100})
            st.subheader('Performance comparison')
            chart_metric = st.selectbox('Compare by', ['DB', 'PQM GRT', 'FTNR %'])
            ascending = chart_metric == 'FTNR %'
            top = chart_df.sort_values(chart_metric, ascending=ascending).head(10)
            fig = px.bar(top, x=chart_metric, y='Entity', orientation='h', color=chart_metric, color_continuous_scale='Tealgrn' if ascending else 'Blues', title=('Lowest 10' if ascending else 'Top 10') + ' · ' + chart_metric)
            fig.update_layout(yaxis={'categoryorder':'total ascending'}, height=440, margin=dict(l=10,r=10,t=55,b=20), paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
            st.caption('Charts use the selected summary and filters. Lower FTNR is better.')
    else:
        cards = ['<div class="kpi-grid">']
        for group, specs in groups:
            for label, off in specs:
                cards.append(f'<div class="kpi-tile"><div class="kpi-group-label">{escape(group)}</div><div class="kpi-label">{escape(label)}</div><div class="kpi-number">{escape(fmt(metric(off),off))}</div></div>')
        cards.append('</div>')
        st.markdown(''.join(cards), unsafe_allow_html=True)
    st.caption(f'{len(filtered):,} branch detail rows · Scroll horizontally to see all metrics.' if branch_detail else f'{len(filtered):,} of {len(df):,} rows · Scroll horizontally to see all metrics.')
    if filtered.empty:
        st.info('No matching rows. Clear the filters.')
    elif view == '📋 Table View':
        st.markdown(render_table(info, filtered.values.tolist(), info['total'] if all_rows else None), unsafe_allow_html=True)
    st.caption('Figures come directly from the selected Excel summary. The Excel total is shown only for the full view. Filtered percentages are not averaged.')



if __name__ == '__main__':
    main()
