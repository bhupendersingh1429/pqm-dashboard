"""PQM summary dashboard. Run: streamlit run pq_dash.py"""
from pathlib import Path
from io import BytesIO
from html import escape
import hashlib
import hmac
import os
from numbers import Real
import posixpath
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
SHEETS = {'Zone': ('Zone Summary', 3, 37),
          'Region(AVP)': ('Region(AVP) Summary', 4, 38),
          'DM': ('DM Summary', 4, 38),
          'PQM': ('PQM Summary', 8, 42)}
BASE = Path(__file__).resolve().parent


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
                if r2 <= 3 and c1 <= width:
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
    if offset in (19, 27, 33):
        return f'{value:.2%}'
    if offset in (4, 7, 11, 18, 23, 28, 29, 30, 31, 32):
        return f'{value:,.2f}'
    if offset == 24:
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


def atomic_save(path, raw):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix='.tmp')
    try:
        with os.fdopen(fd, 'wb') as f:
            f.write(raw)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


def main():
    import streamlit as st
    st.set_page_config(page_title='PQM Dashboard', layout='wide')
    st.markdown('''<style>
    .stApp{background:#f4f7fb}.block-container{padding-top:2rem}
    .hero{background:linear-gradient(100deg,#102b50,#0b7496);color:white;padding:24px 30px;border-radius:16px;margin-bottom:20px}
    .kpi-group-title{font-size:15px;font-weight:700;color:#123956;margin-top:12px}.kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(175px,1fr));gap:12px;margin:14px 0 20px}.kpi-tile{background:white;border:1px solid #dce5ef;border-top:3px solid #167b9d;border-radius:10px;padding:13px 15px}.kpi-label{font-size:12px;color:#536579;min-height:32px}.kpi-number{font-size:23px;font-weight:700;color:#123956;white-space:nowrap}
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
    st.markdown('<div class="hero"><h1>PQM Dashboard</h1><p>Summary reports · Process &amp; Quality</p></div>', unsafe_allow_html=True)
    def setting(name):
        try:
            return st.secrets.get(name, os.environ.get(name, ''))
        except (FileNotFoundError, st.errors.StreamlitSecretNotFoundError):
            return os.environ.get(name, '')
    directory = setting('PQM_REPORT_DIR')
    shared_path = Path(directory) / 'Report.xlsx' if directory else None
    source = shared_path if shared_path and shared_path.exists() else BASE / 'Report.xlsx'
    st.sidebar.title('Report controls')
    level = st.sidebar.selectbox('Summary', list(SHEETS))
    if st.sidebar.button('Refresh report'):
        st.rerun()
    password = setting('PQM_ADMIN_PASSWORD')
    with st.sidebar.expander('Upload latest report'):
        st.caption('An upload preview is visible only in your current session.')
        upload = st.file_uploader('Combined PQM Report', type=['xlsx'])
        if upload is not None and st.button('Preview uploaded report'):
            raw = upload.getvalue()
            try:
                load_report(raw)
                st.session_state['preview'] = raw
                st.session_state['preview_name'] = upload.name
                st.rerun()
            except Exception as e:
                st.error(f'Cannot read report: {e}')
        if 'preview' in st.session_state:
            if password and shared_path:
                token = st.text_input('Admin password', type='password')
                if st.button('Publish report for all users'):
                    if hmac.compare_digest(token, str(password)):
                        try:
                            atomic_save(shared_path, st.session_state['preview'])
                            del st.session_state['preview']
                            st.success('Shared report updated.')
                            st.rerun()
                        except OSError as e:
                            st.error(f'Could not save the shared report: {e}')
                    else:
                        st.error('Incorrect password.')
            else:
                st.info('Shared publishing needs an admin password and a persistent report folder. You can also replace Report.xlsx in the deployed repository.')
            if st.button('Return to shared report'):
                del st.session_state['preview']
                st.rerun()
    try:
        raw = st.session_state.get('preview')
        if raw is None:
            raw = source.read_bytes()
        cached_load = st.cache_data(show_spinner='Reading summaries…')(load_report)
        reports = cached_load(raw)
    except Exception as e:
        st.error(f'Could not open report: {e}')
        st.info('Place the new Combined PQM Report next to pq_dash.py, named Report.xlsx.')
        st.stop()
    info = reports[level]
    st.subheader(info['name'])
    st.caption(info['date'])
    if 'preview' in st.session_state:
        st.warning('Upload preview — other users are still viewing the shared report.')
    else:
        st.caption(f'Shared report · {hashlib.sha256(raw).hexdigest()[:8]}')
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
    if level in ('Zone', 'Region(AVP)'):
        pq_info = reports['PQM']
        branch_rows = pd.DataFrame(pq_info['rows'], columns=pq_info['headers'])
        for field, chosen in entity_filters.items():
            if field in branch_rows.columns:
                branch_rows = branch_rows[branch_rows[field].fillna('').astype(str).isin(chosen)]
        branches = sorted(x for x in branch_rows['Branch'].dropna().astype(str).unique() if x.strip() and x != '-')
        selected_branches = st.sidebar.multiselect('PQ Branch', branches, key=f'{level}-branch')
        if selected_branches or (level == 'Region(AVP)' and entity_filters.get('Region AVP')):
            branch_rows = branch_rows[branch_rows['Branch'].fillna('').astype(str).str.strip().ne('') & branch_rows['Branch'].fillna('').astype(str).str.strip().ne('-')]
            if pq_info['missing']:
                st.error('PQM formula results are missing. Recalculate and save the Excel report first.')
                st.stop()
            info = pq_info
            filtered = (branch_rows[branch_rows['Branch'].fillna('').astype(str).isin(selected_branches)] if selected_branches else branch_rows).copy()
            branch_detail = True
            st.caption('PQ branch performance for the selected Zone / Region · Source: PQM Summary. Clear Zone, Region and PQ Branch filters to return to the summary view.')
    all_rows = not branch_detail and len(filtered) == len(df)
    def sum_col(offset):
        s = pd.to_numeric(filtered.iloc[:, info['ids']+offset], errors='coerce')
        return s.sum() if len(s) and s.notna().all() else None
    # Counts and amounts are additive; rates are ratios of matching totals.
    groups = [
        ('Vishwas', [('Total Allocated', 0), ('Total Conducted', 1), ('DB', 10), ('DB Amount (Cr)', 11), ('Declined', 12)]),
        ('GRT', [('PQM GRT', 16), ('GRT % PQM', 19)]),
        ('Ujala', [('CB to DB', 20), ('DB Amount (Cr)', 24)]),
        ('FTNR', [('FTNR %', 27)]),
        ('Time in Minutes', [('Effective Productivity', 33)])]
    def effective_productivity():
        if filtered.empty:
            return None
        if all_rows and info['total'] is not None:
            value = info['total'][info['ids']+33]
            return value if isinstance(value, Real) else None
        values = pd.to_numeric(filtered.iloc[:, info['ids']+33], errors='coerce')
        if len(filtered) == 1:
            return values.iloc[0] if values.notna().all() else None
        if info['name'] == 'PQM Summary':
            # Each populated employee row is one PQM; exclude unassigned rows.
            weights = filtered['PQM EMP ID'].fillna('').astype(str).str.strip().map(lambda x: 1 if x and x != '-' else 0)
        else:
            count_col = next(c for c in filtered.columns if c.lower().startswith('no') and 'pqm' in c.lower())
            weights = pd.to_numeric(filtered[count_col], errors='coerce').fillna(0)
        valid = weights > 0
        if not valid.any() or values[valid].isna().any():
            return None
        return (values[valid] * weights[valid]).sum() / weights[valid].sum()
    cards = []
    for group, specs in groups:
        cards.append(f'<div class="kpi-group-title">{escape(group)}</div><div class="kpi-grid">')
        for label, off in specs:
            value = effective_productivity() if off == 33 else sum_col(off)
            if off in (19, 27):
                numerator, denominator = (sum_col(16), sum_col(17)) if off == 19 else (sum_col(25), sum_col(26))
                value = numerator / denominator if numerator is not None and denominator else None
            if value is None:
                formatted = '—'
            elif off == 24:
                formatted = f'₹ {value / 10000000:,.2f} Cr'
            elif off == 11:
                formatted = f'₹ {value:,.2f} Cr'
            else:
                formatted = display(value, off)
            cards.append(f'<div class="kpi-tile"><div class="kpi-label">{escape(label)}</div><div class="kpi-number">{escape(formatted)}</div></div>')
        cards.append('</div>')
    st.markdown(''.join(cards), unsafe_allow_html=True)
    st.caption(f'{len(filtered):,} branch detail rows · Scroll horizontally to see all metrics.' if branch_detail else f'{len(filtered):,} of {len(df):,} rows · Scroll horizontally to see all metrics.')
    if filtered.empty:
        st.info('No matching rows. Clear the filters.')
    else:
        st.markdown(render_table(info, filtered.values.tolist(), info['total'] if all_rows else None), unsafe_allow_html=True)
    st.caption('Figures come directly from the selected Excel summary. The Excel total is shown only for the full view. Filtered percentages are not averaged.')
    st.download_button('Download selected rows (CSV)', filtered.to_csv(index=False).encode('utf-8-sig'),
                       file_name=f'PQM_{level.replace("(AVP)", "")}_summary.csv', mime='text/csv')
    st.download_button('Download full Excel report', raw, file_name='Combined_PQM_Report.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    main()
