"""Run: streamlit run pq_dashboard_updated.py
Install: pip install streamlit pandas plotly pyxlsb openpyxl
Source: latest report layout (Branch Summary A:AN, header in row 4).
Upload .xlsb/.xlsx or place Report.xlsb / Report.xlsx beside this script.
"""
from pathlib import Path
from io import BytesIO
import re
import pandas as pd

COLS = [
    'Zone','State','RM','DM','PQM','Employee ID','Branch',
    'Branch GRT','PQM GRT','Total GRT','GRT Daily Source','GRT % Source',
    'Ujala CB to DB','Last Day STH','Last Day DB','Ujala Daily Source','Ujala DB Amount',
    'Physical PD','Last Day Physical PD','PD Daily Source','Video PD','Last Day Video PD',
    'VPD Daily Source','Branch FTNR','Branch FTNR Base','Branch FTNR % Source',
    'PQM FTNR','PQM FTNR Base','PQM FTNR % Source',
    'Same Day Source','T+1 Source','T+2 Source','>2 Source',
    'GRT Minutes','Ujala Minutes','PD Minutes','VPD Minutes','Total Minutes',
    'Effective Productivity Source','Designation']
TEXT = COLS[:7] + ['Designation']
ADDITIVE = ['Branch GRT','PQM GRT','Total GRT','Ujala CB to DB','Last Day STH',
    'Last Day DB','Ujala DB Amount','Physical PD','Last Day Physical PD',
    'Video PD','Last Day Video PD','Branch FTNR','Branch FTNR Base',
    'PQM FTNR','PQM FTNR Base','Physical PD Declined','Video PD Declined']
TAT = ['Same Day','T+1','T+2','>2']

def text(value):
    return '' if pd.isna(value) else str(value).strip()

def engine_for(name):
    return 'pyxlsb' if name.lower().endswith('.xlsb') else 'openpyxl'

def strict_sum(s):
    # Missing/error values are unknown, never silently counted as zero.
    return s.sum() if len(s) and s.notna().all() else float('nan')

def ratio(a,b):
    return a / b * 100 if pd.notna(a) and pd.notna(b) and b > 0 else float('nan')

def load_report(raw, name):
    book = pd.ExcelFile(BytesIO(raw), engine=engine_for(name))
    source = pd.read_excel(book, sheet_name='Branch Summary', header=None)
    if source.shape[1] < 40 or source.shape[0] < 5:
        raise ValueError('Branch Summary needs the latest 40-column layout with headers in row 4.')
    headers = [text(x) for x in source.iloc[3,:7]]
    if headers != ['Zone','State','RM Name','DM Name','PQM Name','PQM Emp ID','Branch']:
        raise ValueError('Unexpected Branch Summary headers. Expected Zone, State, RM Name, DM Name, PQM Name, PQM Emp ID, Branch in row 4.')
    df = source.iloc[4:,:40].copy()
    df.columns = COLS
    df = df[df['Branch'].notna() & df['Zone'].notna()].copy()
    for c in TEXT:
        df[c] = df[c].map(text)
    df = df[~df['Zone'].str.casefold().eq('total')].copy()
    for c in set(COLS)-set(TEXT):
        df[c] = pd.to_numeric(df[c], errors='coerce')
    df['Team'] = df['DM'].str.casefold().eq('central pq').map({True:'Central PQ',False:'Field PQ'})
    df.loc[df['Team'].eq('Central PQ'),'RM'] = 'Central PQ'
    df['RM'] = df['RM'].replace({'':'Unmapped RM','-':'Unmapped RM'})
    df['DM'] = df['DM'].replace({'':'Unmapped DM','-':'Unmapped DM'})
    if df['Employee ID'].duplicated().any() or df['Employee ID'].eq('').any():
        raise ValueError('Missing or duplicate employee IDs in Branch Summary. Resolve mapping before loading.')
    notes = []
    # Count performed/declined cases by actual employee, not the customer branch.
    # This retains Video PD work on non-PQ customer branches.
    for sheet, id_col, metric, last, decline in [
        ('PD','PD done ID','Physical PD','Last Day Physical PD','Physical PD Declined'),
        ('Video PD','Video PD done ID','Video PD','Last Day Video PD','Video PD Declined')]:
        if sheet not in book.sheet_names:
            df[decline] = float('nan')
            notes.append(f'{sheet} detail sheet missing; decline counts unavailable.')
            continue
        detail = pd.read_excel(book,sheet_name=sheet)
        detail.columns = [text(c) for c in detail.columns]
        required = [id_col,'DM','Done by',"Today's or not",'Declined Status']
        if not set(required).issubset(detail.columns):
            df[decline] = float('nan')
            notes.append(f'{sheet} detail headers incomplete; using summary counts.')
            continue
        detail = detail[detail['Customer Number'].notna()].copy() if 'Customer Number' in detail else detail
        detail['_id'] = detail[id_col].map(text).str.upper()
        detail['_dm'] = detail['DM'].map(text)
        known = set(df['Employee ID'].str.upper())
        unknown = ~detail['_id'].isin(known)
        # Keep unmapped work visible, scoped to its DM. No invented RM allocation.
        for dm in detail.loc[unknown,'_dm'].unique():
            part = unknown & detail['_dm'].eq(dm)
            key = ('UNMAPPED::' + (dm or 'Unknown DM')).upper()
            detail.loc[part,'_id'] = key
            if key not in set(df['Employee ID']):
                same_dm = df[df['DM'].eq(dm)]
                rms = same_dm['RM'].unique()
                central = dm.casefold() == 'central pq'
                row = {c:float('nan') for c in COLS if c not in TEXT}
                row.update({c:'' for c in TEXT})
                row.update({c:0 for c in ['Physical PD','Last Day Physical PD','Physical PD Declined','Video PD','Last Day Video PD','Video PD Declined']})
                row.update({'RM':'Central PQ' if central else (rms[0] if len(rms)==1 else 'Unmapped RM'),
                    'DM':dm or 'Unknown DM','PQM':'Unmapped activity','Employee ID':key,
                    'Branch':'Unmapped activity','Team':'Central PQ' if central else 'Field PQ',
                    'Designation':'Unmapped'})
                df = pd.concat([df,pd.DataFrame([row])],ignore_index=True)
        counts = detail.groupby('_id').size()
        daily = detail[detail["Today's or not"].map(text).str.casefold().eq('yes')].groupby('_id').size()
        declined = detail[detail['Declined Status'].map(text).str.casefold().eq('declined')].groupby('_id').size()
        previous = df[metric].sum()
        df[metric] = df['Employee ID'].str.upper().map(counts).fillna(0)
        df[last] = df['Employee ID'].str.upper().map(daily).fillna(0)
        df[decline] = df['Employee ID'].str.upper().map(declined).fillna(0)
        if previous != len(detail):
            notes.append(f'{metric}: Branch Summary {previous:,.0f}; detail {len(detail):,}. Dashboard uses detail counts; unmapped activity remains visible.')
    # Synthetic unmapped rows have no summary activities beyond PD/VPD.
    synthetic = df['Employee ID'].str.startswith('UNMAPPED::')
    for c in ['Branch GRT','PQM GRT','Total GRT','Ujala CB to DB','Last Day STH','Last Day DB','Ujala DB Amount','PQM FTNR','PQM FTNR Base']:
        df.loc[synthetic,c] = 0
    asof = 'Date unavailable'
    for value in source.iloc[:4].values.ravel():
        match = re.search(r'As of\s+(\d{1,2}-[A-Za-z]{3}-\d{4})',str(value),re.I)
        if match:
            asof = match.group(1)
            break
    days = pd.to_numeric(pd.Series([source.iat[1,41] if source.shape[1]>41 else None]),errors='coerce').iloc[0]
    if not pd.notna(days) or days <= 0:
        days = None
        notes.append('Working days unavailable; daily productivity needs a valid working-day value.')
    if df['Branch FTNR'].isna().any():
        notes.append('Branch FTNR contains missing/error values. Affected totals and percentages display N/A.')
    if 'Ujala Summary' in book.sheet_names:
        uj = pd.read_excel(book,sheet_name='Ujala Summary',header=None,nrows=3)
        uj_amount = pd.to_numeric(pd.Series([uj.iat[1,7]]),errors='coerce').iloc[0]
        summary_amount = df['Ujala DB Amount'].sum()
        if pd.notna(uj_amount) and uj_amount != summary_amount:
            notes.append(f'Ujala amount differs: Branch Summary Rs {summary_amount:,.0f}; Ujala Summary Rs {uj_amount:,.0f}. Dashboard uses Branch Summary, labelled accordingly.')
    return df, asof, days, notes

def aggregate(df, level, days):
    keys = {'RM':['RM'],'DM':['RM','DM'],'Branch':['RM','DM','Branch']}[level]
    records=[]
    for key,g in df.groupby(keys,dropna=False,sort=True):
        key = key if isinstance(key,tuple) else (key,)
        row=dict(zip(keys,key))
        staff=g[~g['Employee ID'].str.startswith('UNMAPPED::')]
        field=staff[staff['Team'].eq('Field PQ')]
        row['Branches']=field['Branch'].nunique()
        row['Staff']=staff['Employee ID'].nunique()
        for c in ADDITIVE:
            scope = g if c in ['Physical PD','Last Day Physical PD','Video PD','Last Day Video PD','Physical PD Declined','Video PD Declined'] else g[g['Team'].eq('Field PQ')]
            row[c]=strict_sum(scope[c]) if c in scope else float('nan')
        row['PQM GRT %']=ratio(row['PQM GRT'],row['Total GRT'])
        row['PQM FTNR %']=ratio(row['PQM FTNR'],row['PQM FTNR Base'])
        row['Branch FTNR %']=ratio(row['Branch FTNR'],row['Branch FTNR Base'])
        row['Physical PD Decline %']=ratio(row['Physical PD Declined'],row['Physical PD'])
        row['Video PD Decline %']=ratio(row['Video PD Declined'],row['Video PD'])
        for c in ['PQM GRT','Ujala CB to DB','Physical PD','Video PD']:
            row[c+' / Staff / Day'] = row[c]/row['Staff']/days if row['Staff'] and days else float('nan')
        # Reconstruct TAT bucket counts, weighting by each row's PQM GRT.
        active=staff[staff['PQM GRT'].gt(0)]
        for c in TAT:
            row[c+' %']=ratio(strict_sum(active[c+' Source']*active['PQM GRT']),strict_sum(active['PQM GRT']))
        for c in ['GRT Minutes','Ujala Minutes','PD Minutes','VPD Minutes','Total Minutes']:
            row[c+' / Staff']=strict_sum(staff[c])/len(staff) if len(staff) else float('nan')
        row['Effective Productivity %']=strict_sum(staff['Effective Productivity Source'])/len(staff)*100 if len(staff) else float('nan')
        records.append(row)
    result=pd.DataFrame(records)
    if not result.empty:
        result.insert(0,'Name',result[level].astype(str))
    return result

def main():
    import streamlit as st
    import plotly.express as px
    st.set_page_config(page_title='PQ Performance',page_icon='📊',layout='wide')
    st.markdown('''<style>.stApp{background:#f3f6fb} [data-testid="stMetric"]{background:white;padding:16px;border-radius:12px;border-top:4px solid #1967ad} h1,h2,h3{color:#103968}</style>''',unsafe_allow_html=True)
    st.title('PQ Performance Dashboard')
    st.sidebar.header('Report & Filters')
    upload=st.sidebar.file_uploader('Latest report',type=['xlsb','xlsx'])
    if upload:
        raw,name=upload.getvalue(),upload.name
    else:
        candidates=[Path(__file__).parent/n for n in ['Report.xlsb','Report.xlsx']]
        local=next((p for p in candidates if p.exists()),None)
        if local is None:
            st.info('Upload your latest .xlsb or .xlsx report from the sidebar to begin.')
            st.stop()
        raw,name=local.read_bytes(),local.name
    try:
        data,asof,report_days,notes=st.cache_data(show_spinner='Reading report...')(load_report)(raw,name)
    except Exception as exc:
        st.error(f'Could not read report: {exc}')
        st.stop()
    st.caption(f'As of {asof} • {name}')
    days=st.sidebar.number_input('Working days',min_value=1,value=int(report_days or 1))
    if report_days is None:
        st.warning('Working days were not found. Set the correct value in the sidebar before using productivity.')
    team=st.sidebar.radio('Team',['Field PQ','Central PQ','All Teams'])
    f=data.copy() if team=='All Teams' else data[data['Team'].eq(team)].copy()
    def select(column,label):
        nonlocal f
        options=sorted(f[column].dropna().unique())
        chosen=st.sidebar.multiselect(label,options,key=team+'_'+column)
        if chosen:
            f=f[f[column].isin(chosen)]
        return chosen
    rms=select('RM','RM Name')
    dms=select('DM','DM Name')
    branches=select('Branch','Branch')
    if f.empty:
        st.info('No records match these filters.')
        st.stop()
    default='Branch' if dms or branches else ('DM' if rms else 'RM')
    level=st.sidebar.selectbox('Comparison level',['RM','DM','Branch'],index=['RM','DM','Branch'].index(default))
    summary=aggregate(f,level,days)
    total=aggregate(f.assign(RM='Selection'),'RM',days).iloc[0]
    def fmt(v,kind='number'):
        if pd.isna(v): return 'N/A'
        if kind=='amount':return f'₹ {v/10000000:.2f} Cr'
        if kind=='percent':return f'{v:.2f}%'
        return f'{v:,.0f}'
    def cards(items):
        for col,(label,metric,kind) in zip(st.columns(len(items)),items):
            col.metric(label,fmt(total[metric],kind))
    cards([('Branches','Branches','number'),('Total GRT','Total GRT','number'),('PQM GRT','PQM GRT','number'),('PQM GRT %','PQM GRT %','percent')])
    cards([('Ujala CB to DB','Ujala CB to DB','number'),('Ujala Amount*','Ujala DB Amount','amount'),('Physical PD','Physical PD','number'),('Video PD','Video PD','number'),('PQM FTNR %','PQM FTNR %','percent')])
    tabs=st.tabs(['Overview','GRT & TAT','Ujala','Physical & Video PD','Quality','Productivity','Summary & Checks'])
    def chart(metric,title,container=st):
        valid=summary.dropna(subset=[metric]).sort_values(metric)
        if valid.empty:
            container.info(f'{title}: no valid data.')
            return
        fig=px.bar(valid,x=metric,y='Name',orientation='h',text_auto='.2f' if '%' in metric or '/ Day' in metric else ',.0f',title=title,color_discrete_sequence=['#1967ad'])
        fig.update_layout(height=max(360,len(valid)*30+100),yaxis_title='',paper_bgcolor='white',plot_bgcolor='white')
        container.plotly_chart(fig,use_container_width=True)
    with tabs[0]:
        chart('PQM GRT %',f'PQM GRT Contribution by {level}')
        chart('Physical PD',f'Physical PD by {level}')
        chart('Video PD',f'Video PD by {level}')
    with tabs[1]:
        chart('PQM GRT',f'PQM GRT by {level}')
        valid=summary.dropna(subset=[c+' %' for c in TAT])
        if not valid.empty:
            fig=px.bar(valid,y='Name',x=[c+' %' for c in TAT],orientation='h',barmode='stack',title='CGT to GRT TAT — weighted by PQM GRT')
            fig.update_layout(height=max(360,len(valid)*30+100),yaxis_title='')
            st.plotly_chart(fig,use_container_width=True)
    with tabs[2]:
        st.caption('*Ujala amount and CB-to-DB count use Branch Summary. They are not relabelled as total company Ujala disbursements.')
        for metric in ['Ujala CB to DB','Ujala DB Amount','Last Day STH','Last Day DB']:
            chart(metric,f'{metric} by {level}')
    with tabs[3]:
        st.caption('Performed and declined counts come from PD/Video PD detail sheets where available. Declined cases are already included in total cases; do not add them again.')
        for metric in ['Last Day Physical PD','Last Day Video PD','Physical PD Declined','Video PD Declined']:
            chart(metric,f'{metric} by {level}')
    with tabs[4]:
        chart('PQM FTNR %',f'PQM FTNR % by {level} — lower is better')
        chart('Branch FTNR %',f'Branch FTNR % by {level}')
        st.caption('PQM FTNR scope: Regular + Ujala + Physical PD. Ratios are recalculated from counts; unknown source counts remain N/A.')
    with tabs[5]:
        st.caption('Daily rates = activity count / mapped staff / working days, including mapped staff with zero activity. Unmapped activity has no extra staff denominator. Time and effective productivity are averages of report values across mapped staff; incomplete values remain N/A. These are report estimates, not measured attendance hours.')
        for metric in ['PQM GRT / Staff / Day','Ujala CB to DB / Staff / Day','Physical PD / Staff / Day','Video PD / Staff / Day','Effective Productivity %']:
            chart(metric,f'{metric} by {level}')
    with tabs[6]:
        st.subheader(f'{level}-wise Summary')
        st.dataframe(summary.round(2),use_container_width=True,hide_index=True)
        st.download_button('Download current summary (CSV)',summary.to_csv(index=False).encode('utf-8-sig'),file_name=f'PQ_{level}_Summary.csv',mime='text/csv')
        st.subheader('Source reconciliation')
        if notes:
            for note in notes: st.warning(note)
        else:st.success('No source discrepancies detected by the available checks.')
        st.caption('Central PQ staff are excluded from physical branch counts. Unmapped activity stays visible under its known DM; it is never assigned to an individual by guesswork. This file is a current snapshot; monthly comparisons require historical reports.')
        with st.expander('Filtered source rows'):
            st.dataframe(f,use_container_width=True,hide_index=True)
    st.subheader(f'Top & Bottom {level} Performance')
    metric=st.selectbox('Ranking metric',['PQM GRT %','PQM FTNR %','Ujala CB to DB','Physical PD','Video PD','Effective Productivity %'])
    eligible=summary.dropna(subset=[metric]).copy()
    if metric in ['Physical PD','Video PD','Effective Productivity %']:
        eligible=eligible[eligible['Staff'].gt(0)]
    eligible=eligible.sort_values([metric,'Name'],ascending=[metric=='PQM FTNR %',True])
    count=min(10,len(eligible))
    if len(eligible)<20:
        st.caption('Fewer than 20 groups: Top and Bottom lists may overlap.')
    a,b=st.columns(2)
    display=['Name','Staff','Branches',metric]
    a.markdown(f'**Top {count}**')
    a.dataframe(eligible.head(count)[display].round(2),hide_index=True,use_container_width=True)
    b.markdown(f'**Bottom {count}**')
    b.dataframe(eligible.tail(count).iloc[::-1][display].round(2),hide_index=True,use_container_width=True)

if __name__=='__main__':
    main()
