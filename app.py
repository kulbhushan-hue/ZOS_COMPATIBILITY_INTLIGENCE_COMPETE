"""z/OS Compatibility Intelligence — seven-page Streamlit application.

Run: streamlit run app.py
Data provenance: original six vendor CSVs (baseline); separately stored Broadcom vendor snapshots.
Refresh executes a source read on demand. Failed scrapes never remove the preceding snapshot.
"""
from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st

from compatibility_db import CompatibilityDB, VENDORS, ZOS_VERSIONS, VER, SOURCES, build_excel, to_output
from crawler import BROADCOM_URL, crawl_broadcom
from store import DEFAULT_DB, latest, changes, jobs, save, record_error
from rag_engine import RagIndex, generate_answer, watsonx_configured

BASE=Path(__file__).resolve().parent
DB_PATH=Path(os.getenv('ZOS_DB_PATH',str(DEFAULT_DB)))
XLSX_MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
PAGE_NAMES=['Overview','Compatibility Matrix','PTF Intelligence','Source Monitor','Change History','Inventory Readiness','AI Assistant']
st.set_page_config(page_title='z/OS Compatibility Intelligence', page_icon='🖥️',layout='wide',initial_sidebar_state='expanded')
st.markdown('''<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');
html,body,[class*="st-"]{font-family:"IBM Plex Sans",Arial,sans-serif}
.block-container{padding-top:1.3rem;max-width:1550px}
[data-testid="stSidebar"]{border-right:1px solid #d1d5db}
[data-testid="stMetric"]{padding:.9rem;border:1px solid rgba(128,128,128,.25);border-radius:10px}
.brand{padding:20px 24px;background:linear-gradient(110deg,#102642,#20466f);border-radius:12px;color:white;margin-bottom:18px}
.brand h1{font-size:1.8rem;color:white;margin:0}.brand p{color:#d5e4f8;margin:7px 0 0}
.small-note{font-size:12px;color:#65738a}
[data-testid="stDataFrame"]{border-radius:8px}
</style>''',unsafe_allow_html=True)

@st.cache_resource
def base_db():
    return CompatibilityDB.load(str(BASE/'data'))


def source_records():
    try:return latest(path=DB_PATH)
    except Exception as exc:
        st.warning(f'Snapshot store not readable: {exc}')
        return [],None


def combined_data(db:CompatibilityDB,live:list[dict])->pd.DataFrame:
    """Keep CSV baseline intact; replace only exactly matched vendor/product/VRM+OS cells."""
    frame=db.df.copy()
    for ver in ZOS_VERSIONS:
        frame[f'provenance_{ver}']='CSV baseline'
    frame['live_details_count']=0
    if not live:return frame
    # Merge conservative exact matches. Do not accidentally map two vendor releases.
    def norm_name(x):return re.sub(r'[^a-z0-9]+','',str(x).lower())
    index={(norm_name(r.vendor),norm_name(r.product),norm_name(r.vrm)):i for i,r in frame.iterrows()}
    added=[]
    for r in live:
        if r['zos_version'] not in VER:continue
        col,ptf_col,*_=VER[r['zos_version']]
        k=(norm_name(r['vendor']),norm_name(r['product']),norm_name(r['vrm']))
        index_number=index.get(k)
        if index_number is None:
            # Include scraped rows that are absent from the original CSV.
            if k not in added:
                record={c:'' for c in frame.columns}
                record.update({'vendor':r['vendor'],'product':r['product'],'vrm':r['vrm'],'label':f"{r['vendor']} | {r['product']} | {r['vrm']}",
                               'link':r['source_url'],'data_status':'Live vendor snapshot','last_checked':'Snapshot'})
                for v in ZOS_VERSIONS:record[VER[v][0]]='Unverified';record[VER[v][1]]='Verify';record[f'provenance_{v}']='No baseline'
                frame=pd.concat([frame,pd.DataFrame([record])],ignore_index=True)
                index_number=len(frame)-1;index[k]=index_number;added.append(k)
            else:index_number=index[k]
        frame.at[index_number,col]=r['status']
        frame.at[index_number,f'provenance_{r["zos_version"]}']='Vendor snapshot'
        if r['ptf_links_present']:
            ids=', '.join(r['ptfs']) if r['ptfs'] else 'Open PTF link; exact identifiers not extracted'
            frame.at[index_number,ptf_col]='Yes - '+ids
        elif r['status']!='Yes':frame.at[index_number,ptf_col]='Verify'
        frame.at[index_number,'live_details_count']=int(frame.at[index_number,'live_details_count'] or 0)+len(r['links'])
    return frame


def status_counts(d,ver):
    c=VER[ver][0]
    return d[c].value_counts().reindex(['Yes','No','In Progress','Unverified','N/A','Unknown'],fill_value=0)


def format_links(r):
    links=r.get('links',[])
    ptfs=', '.join(r.get('ptfs',[]))
    return ptfs or ('PTF link available; identifiers not extracted' if r.get('ptf_links_present') else 'No PTF link in this cell')


def detail_frame(rows):
    return pd.DataFrame([{'Type':l['type'],'PTF IDs':', '.join(l['ptfs']),'APAR IDs':', '.join(l['apars']),
        'FIXCAT':', '.join(l['fixcats']),'USERMOD':', '.join(l.get('usermods',[])),'SYSMOD':', '.join(l.get('sysmods',[])),'FMIDs':', '.join(l.get('fmids',[])),'Extraction':l['extraction_status'],'Source URL':l['url'],'Extracted text':l['text']}
        for r in rows for l in r.get('links',[])])


def excel_bytes(data):
    buf=BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as w:
        for name,df in data.items():
            copy=df.copy()
            for col in copy.columns:
                copy[col]=copy[col].map(lambda v:json.dumps(v) if isinstance(v,(dict,list)) else v)
            copy.to_excel(w,sheet_name=name[:31],index=False)
            ws=w.sheets[name[:31]]
            ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=cell.font.copy(bold=True)
            for col in ws.columns:
                letter=col[0].column_letter
                ws.column_dimensions[letter].width=min(57,max(13,max(len(str(c.value or '')) for c in list(col)[:150])+2))
    return buf.getvalue()


def page_header(title,subtitle):
    st.markdown(f'<div class="brand"><h1>{title}</h1><p>{subtitle}</p></div>',unsafe_allow_html=True)


base=base_db()
records,info=source_records()
all_df=combined_data(base,records)
db=CompatibilityDB(all_df)

with st.sidebar:
    st.markdown('### ◈ z/OS Intelligence')
    page=st.radio('Workspace',PAGE_NAMES,label_visibility='collapsed')
    st.divider()
    zos=st.selectbox('Target z/OS release',ZOS_VERSIONS,key='global_zos')
    st.metric('Vendor records',len(db.df))
    st.metric('Vendors in baseline',base.df['vendor'].nunique())
    st.caption('Broadcom snapshot: '+(str(info['captured_at'])[:19]+' UTC' if info else 'Not yet collected'))
    st.divider()
    st.caption('Source distinction: CSV baseline vs. published vendor snapshot. A compatibility YES is not confirmation of installed PTFs.')


if page=='Overview':
    page_header('Compatibility Overview','Mainframe software intelligence · vendor statements, maintenance and upgrade risks')
    if info:st.success(f'Latest Broadcom snapshot: {info["row_count"]} version-specific records collected {info["captured_at"][:19]} UTC')
    else:st.info('Displaying CSV baseline. To collect current Broadcom maintenance evidence, open Source Monitor and select Refresh Broadcom.')
    stats=db.summarize(db.df,zos)
    cols=st.columns(6)
    for col,(lab,num) in zip(cols,[('Products',stats['total']),('Compatible',stats['yes']),('Not compatible',stats['no']),('In progress',stats['in_progress']),('Unverified',stats['unverified']),('Require PTF review',stats['ptf'])]):col.metric(lab,num)
    a,b=st.columns(2)
    with a:
        st.subheader('Compatibility statuses')
        st.bar_chart(status_counts(db.df,zos).rename('Records'),height=320)
    with b:
        st.subheader('Vendor coverage')
        st.bar_chart(db.df.groupby('vendor').size().rename('Products'),height=320)
    st.subheader('Recent source jobs')
    history=jobs(5,DB_PATH)
    if history:st.dataframe(pd.DataFrame(history),hide_index=True,use_container_width=True)
    else:st.caption('No fetch attempts recorded. The app does not scrape vendors automatically on every page view.')
    st.subheader('Important distinctions')
    st.write('**Compatible:** the vendor indicates support for that z/OS release. **Maintenance requirements:** linked PTFs/APARs/prerequisites may still be needed. **Upgrade ready:** cannot be confirmed until you compare these with installed mainframe maintenance.')

elif page=='Compatibility Matrix':
    page_header('Compatibility Matrix','Browse all six vendor baselines; override matching Broadcom rows with the latest successful snapshot')
    a,b,c,d=st.columns([1,1,1,2])
    vend=a.multiselect('Vendors',VENDORS,default=VENDORS)
    ver=b.selectbox('Target release',ZOS_VERSIONS,index=ZOS_VERSIONS.index(zos))
    status=c.multiselect('Status',['Yes','No','In Progress','Unverified','N/A','Unknown'])
    q=d.text_input('Search product / version / PTF',placeholder='ACF2, Endevor, Db2, LU15603 ...')
    filtered=db.filter(vendors=vend,status=status or None,search=q,version=ver)
    summary=db.summarize(filtered,ver)
    c1,c2,c3,c4=st.columns(4)
    c1.metric('Matching records',len(filtered));c2.metric('Compatible',summary['yes']);c3.metric('Not compatible',summary['no']);c4.metric('Unverified',summary['unverified'])
    fields=['vendor','product','vrm',VER[ver][0],VER[ver][1],f'provenance_{ver}','eos_date','latest_version','link']
    st.dataframe(filtered[fields].rename(columns={'vendor':'Vendor','product':'Product','vrm':'Release',VER[ver][0]:'Compatibility',VER[ver][1]:'PTF required',f'provenance_{ver}':'Evidence','eos_date':'EOS date','latest_version':'Latest version','link':'Vendor source'}),hide_index=True,use_container_width=True,height=450,column_config={'Vendor source':st.column_config.LinkColumn('Vendor source')})
    st.download_button('⬇ Download filtered Excel',excel_bytes({'Compatibility':filtered[fields]}),file_name=f'zos_{ver}_filtered.xlsx',mime=XLSX_MIME)
    if records:
        with st.expander('View scraped-only matrix rows / maintenance details'):
            rows=[r for r in records if r['zos_version']==ver and r['vendor'] in vend]
            st.dataframe(pd.DataFrame([{'Product':r['display_name'],'Status':r['status'],'PTFs':', '.join(r['ptfs']),'PTF links':r['ptf_links_present'],'Source':r['source_url']} for r in rows]),use_container_width=True,hide_index=True)

elif page=='PTF Intelligence':
    page_header('PTF Intelligence','Exact PTF identifiers, linked instructions, fix categories and traceable evidence')
    if not records:
        st.info('No live maintenance snapshot exists yet. Open Source Monitor → Refresh Broadcom to fetch source PTF and Additional Information details.')
        st.markdown('The existing CSV baseline still provides preliminary PTF flags in the Compatibility Matrix.')
    else:
        options=sorted({(r['product'],r['vrm']) for r in records})
        pair=st.selectbox('Product / release',options,format_func=lambda t:f'{t[0]} {t[1]}')
        version=st.selectbox('Target z/OS',ZOS_VERSIONS,index=ZOS_VERSIONS.index(zos))
        sel=[r for r in records if (r['product'],r['vrm'])==pair and r['zos_version']==version]
        if sel:
            r=sel[0]
            a,b,c,d=st.columns(4)
            a.metric('Compatibility',r['status']);b.metric('PTFs identified',len(r['ptfs']))
            c.metric('Detail links',len(r['links']));d.metric('Evidence',r['evidence_status'])
            t1,t2,t3,t4=st.tabs(['Required PTFs','Additional Information','OS Comparison','Source Evidence'])
            with t1:
                if r['ptfs']:st.code(', '.join(r['ptfs']),language=None)
                elif r['ptf_links_present']:st.warning('PTF link exists, but exact identifiers could not be verified in the extracted text.')
                else:st.info('No PTF link listed for this product and OS. That is not proof that maintenance is unnecessary.')
                st.caption('PTF identifiers must be checked against vendor instructions and installed SMP/E records.')
                df=detail_frame([r]);subset=df[df['Type'].str.contains('PTF',case=False,na=False)] if not df.empty else df
                if not subset.empty:st.dataframe(subset,hide_index=True,use_container_width=True,column_config={'Source URL':st.column_config.LinkColumn('Source URL')})
            with t2:
                df=detail_frame([r]);subset=df[~df['Type'].str.contains('PTF',case=False,na=False)] if not df.empty else df
                if subset.empty:st.info('No linked additional-information content extracted for this combination.')
                else:
                    for _,entry in subset.iterrows():
                        with st.expander(f"{entry['Type']} · {entry['Extraction']}"):
                            st.write(entry['Extracted text'] or 'Link could not be resolved.');st.link_button('Open original reference',entry['Source URL'])
            with t3:
                comparison=[x for x in records if (x['product'],x['vrm'])==pair]
                st.dataframe(pd.DataFrame([{'z/OS':x['zos_version'],'Compatibility':x['status'],'PTFs':', '.join(x['ptfs']),'Detail links':len(x['links']),'Evidence':x['evidence_status']} for x in comparison]),use_container_width=True,hide_index=True)
            with t4:
                st.write('**Source:**',r['source_url']);st.write('**Last snapshot:**',info['captured_at'])
                st.json({'product':r['product'],'vrm':r['vrm'],'zos_version':r['zos_version'],'status':r['status'],'ptfs':r['ptfs'],'evidence_status':r['evidence_status']},expanded=False)
                for link in r['links']:
                    with st.expander(link['type']+' · '+link['extraction_status']):
                        st.write(link['text'] or 'Not readable');st.link_button('Open vendor evidence',link['url'])
            st.download_button('⬇ Download maintenance Excel',excel_bytes({'Details':detail_frame([r]),'Selected record':pd.DataFrame([r]).drop(columns=['links','ptfs'],errors='ignore')}),file_name='zos_maintenance_details.xlsx',mime=XLSX_MIME)

elif page=='Source Monitor':
    page_header('Source Monitor','Controlled collection, vendor evidence coverage, refresh outcomes and detailed logs')
    st.warning('A refresh makes live network requests from the running server. Your first run may encounter rate limits, blocked pages, redirects or changed HTML. The previous successful snapshot will be preserved on failure.')
    st.write('**Broadcom matrix source:**');st.link_button('Open Broadcom matrix',BROADCOM_URL)
    st.caption('The crawler extracts the main product table, resolves internal PTF and Additional Information anchors, and follows approved off-page vendor documents when enabled.')
    fetch_external=st.checkbox('Follow approved external detail pages (slower)',value=False,help='Local anchor links are always expanded. Separate pages can require additional requests.')
    if st.button('↻ Refresh Broadcom now',type='primary'):
        try:
            with st.spinner('Fetching Broadcom matrix, expanding maintenance references, and checking the extracted data...'):
                fresh,stats=crawl_broadcom(follow_external=fetch_external)
                if len(fresh)<30:raise ValueError('Unusually few rows: rejected to avoid publishing incomplete results')
                snap=save(fresh,stats,DB_PATH)
            st.success(f'Snapshot #{snap} published: {stats["records"]} version-specific rows, {stats["links"]} detail links. Open any page to see refreshed data.')
            st.rerun()
        except Exception as exc:
            record_error('Broadcom',str(exc),DB_PATH)
            st.error(f'Refresh failed: {type(exc).__name__}: {exc}. Previous successful data was retained.')
    if info:
        st.write(f'**Last good snapshot:** {info["captured_at"]} UTC · {info["row_count"]} records')
        with st.expander('Extraction statistics'):st.json(json.loads(info['stats_json']))
    source_summary=pd.DataFrame([{'Vendor':v,'Baseline rows':int((base.df['vendor']==v).sum()),'Live source adapter':'Broadcom matrix + linked sections' if v=='Broadcom' else 'Not implemented','Last successful live scrape':info['captured_at'] if v=='Broadcom' and info else 'Never'} for v in VENDORS])
    st.dataframe(source_summary,hide_index=True,use_container_width=True)
    st.subheader('Refresh job history')
    log=jobs(path=DB_PATH)
    if log:st.dataframe(pd.DataFrame(log),hide_index=True,use_container_width=True)
    else:st.caption('No crawl jobs performed yet.')
    st.subheader('Official vendor source directory')
    for v in VENDORS:
        with st.expander(v):
            for vendor,title,url,description in SOURCES:
                if vendor==v:st.markdown(f'[{title}]({url}) — {description}')

elif page=='Change History':
    page_header('Change History','Track changed support statements and linked maintenance across successful source snapshots')
    history=changes(path=DB_PATH)
    if not info:st.info('Collect a first snapshot, then a second successful snapshot. Changes are reported only after comparisons between two real snapshots.')
    elif not history:st.info('No changes recorded. This may be the first snapshot or a repeat with identical information.')
    else:
        st.metric('Recorded changes',len(history))
        table=pd.DataFrame([{'Timestamp':r['captured_at'],'Key':r['row_key'],'Type':r['change_type']} for r in history])
        st.dataframe(table,hide_index=True,use_container_width=True)
        selected=st.selectbox('Inspect change',range(len(history)),format_func=lambda i:history[i]['row_key']+' ('+history[i]['change_type']+')')
        a,b=st.columns(2)
        with a:st.markdown('**Previous**');st.json(json.loads(history[selected]['before_json']) if history[selected]['before_json'] else {},expanded=True)
        with b:st.markdown('**Current**');st.json(json.loads(history[selected]['after_json']) if history[selected]['after_json'] else {},expanded=True)
        st.download_button('⬇ Export change log',excel_bytes({'Changes':table}),file_name='zos_changes.xlsx',mime=XLSX_MIME)

elif page=='Inventory Readiness':
    page_header('Inventory Readiness','Match existing mainframe products to vendor statements and maintenance obligations')
    target=st.selectbox('Target release',ZOS_VERSIONS,index=ZOS_VERSIONS.index(zos))
    up=st.file_uploader('Upload inventory (CSV / Excel: Vendor, Product, VRM)',type=['csv','xlsx'])
    labels=db.df['label'].tolist()
    selected=st.multiselect('Or select individual products',labels)
    chosen=db.df[db.df['label'].isin(selected)]
    misses=[]
    if up:
        try:
            inventory=pd.read_csv(up,dtype=str,keep_default_na=False) if up.name.lower().endswith('.csv') else pd.read_excel(up,dtype=str,keep_default_na=False)
            matched,misses=db.match_inventory(inventory)
            chosen=pd.concat([chosen,matched],ignore_index=True).drop_duplicates(subset=['vendor','product','vrm'])
        except Exception as exc:st.error(f'Inventory import error: {exc}')
    if chosen.empty:st.info('Select products or upload an inventory to analyze compatibility and possible PTF requirements.')
    else:
        stat=db.summarize(chosen,target);cs=st.columns(5)
        for c,(lab,val) in zip(cs,[('Matched',len(chosen)),('Compatible',stat['yes']),('Not compatible',stat['no']),('PTF review',stat['ptf']),('Unverified',stat['unverified'])]):c.metric(lab,val)
        result=to_output(chosen,version=target)
        result['Migration readiness']='Not confirmed — installed PTFs not provided'
        st.dataframe(result,hide_index=True,use_container_width=True,height=400)
        if stat['no']:st.error('One or more products are not compatible with the target z/OS release.')
        if stat['ptf']:st.warning('Vendor-reported PTF requirements must be checked against installed SMP/E service.')
        if stat['unverified'] or stat['in_progress']:st.warning('Some products have no confirmed vendor support statement for the target release.')
        st.download_button('⬇ Download inventory review',excel_bytes({'Assessment':result,'Unmatched':pd.DataFrame({'Issue':misses})}),file_name='zos_inventory_review.xlsx',mime=XLSX_MIME)
    for msg in misses:st.caption('⚠ '+msg)

elif page=='AI Assistant':
    page_header('AI Assistant','Evidence-grounded answers from the compatibility database with optional watsonx.ai')
    st.caption('Without watsonx credentials the application provides retrieval results instead of generating unsupported answers.')
    samples=['What Broadcom products are not compatible with z/OS 3.2?','Does ACF2 16.0 need PTFs?','Which products are still unverified?']
    question=st.text_input('Ask about compatibility or maintenance',placeholder=samples[1])
    k=st.slider('Retrieve matching records',5,30,12)
    if question:
        hits=RagIndex.build(db.df).retrieve(question,k)
        if hits.empty:st.warning('No matching rows found. Try vendor, product or release.')
        else:
            if watsonx_configured():
                try:st.markdown(generate_answer(question,hits))
                except Exception as exc:st.warning('AI answer unavailable; displaying evidence rows instead.')
            else:st.info('Showing matching evidence. Set WATSONX_APIKEY and WATSONX_PROJECT_ID for written answers.')
            st.dataframe(to_output(hits,version=zos),hide_index=True,use_container_width=True)
            if records:
                st.caption('Detailed PTF evidence is available on the PTF Intelligence page. AI text may be incomplete; always check original vendor instructions.')

st.divider()
st.caption('z/OS Compatibility Intelligence · Open vendor sources for authoritative verification · Data is not a guarantee that installed maintenance is sufficient.')
