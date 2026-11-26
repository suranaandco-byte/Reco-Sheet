from pathlib import Path
from io import BytesIO
from calendar import month_name
import re, math
import openpyxl as x
import xlrd

MONTHS=list(month_name)[4:]+list(month_name)[1:4]
TAX=['IGST','CGST','SGST','CESS']
def fill(data, name):
    if name.lower().endswith('.xls'):
        sh=xlrd.open_workbook(file_contents=data).sheet_by_name('Offset Summary')
        get=lambda r,c: sh.cell_value(r-1,c-1) if r<=sh.nrows and c<=sh.ncols else None
    else:
        sh=x.load_workbook(BytesIO(data),data_only=True)['Offset Summary']
        get=lambda r,c: sh.cell(r,c).value
    if str(get(1,2)).strip().lower()!='consolidated 3b offset summary':
        raise ValueError('Expected a Consolidated 3B Offset Summary export.')
    def check(r,c,labels):
        if [str(get(r,c+i)).strip().lower() for i in range(len(labels))]!=[v.lower() for v in labels]:
            raise ValueError(f'Unexpected source layout at row {r}, column {c}.')
    check(6,3,['Outward (NRC)','Outward (RC)','Non-Taxable','Total']+TAX)
    check(5,7,['Liability on Outward']); check(22,7,['Liability on Reverse Charge']); check(23,7,TAX)
    for c,label in zip([4,8,12,16,20],['ITC Eligible (NRC)','ITC Eligible (RC)','ITC Reversed','ITC Ineligible','ITC Reclaimed']):
        check(39,c,[label]); check(40,c,TAX)
    def months(start):
        d={str(get(r,2)).strip().lower():r for r in range(start,start+12)}
        if set(d)!={m.lower() for m in MONTHS}: raise ValueError(f'Expected 12 unique month rows starting at row {start}.')
        return d
    out,rc,itc=map(months,[7,24,41])
    def number(r,c):
        v=get(r,c)
        if v is None or v=='': return None
        if type(v) in (int,float) and math.isfinite(v): return v
        raise ValueError(f'Expected a numeric amount or blank at source row {r}, column {c}.')
    wb=x.load_workbook(Path(__file__).with_name('template.xlsx')); s=wb['Offset Summary']
    for r in (2,3):
        v=get(r,2)
        if not isinstance(v,str) or not v.strip() or v.lstrip().startswith('='): raise ValueError('Missing or invalid company / financial year header.')
        s.cell(r,2).value=v.strip()
    def copy(row,col,sr,sc,n):
        for j in range(n): s.cell(row,col+j).value=number(sr,sc+j)
    for i,m in enumerate(MONTHS):
        k=m.lower(); r=7+i
        copy(r,3,out[k],3,1); copy(r,4,out[k],5,1); copy(r,6,out[k],7,4)
        s[f'E{r}']=f'=IF(COUNT(C{r}:D{r})=0,"",SUM(C{r}:D{r}))'
        s[f'S{r}']=f'=SUM(E{r})-SUM(L{r})'
        copy(25+i,4,itc[k],4,4); copy(42+i,10,rc[k],7,4)
        s.cell(42+i,9).value=m
        copy(60+i,10,itc[k],8,4); copy(60+i,4,itc[k],12,4)
        copy(78+i,7,itc[k],16,4); copy(78+i,2,itc[k],20,4)
        s.cell(78+i,1).value=m
        for d in range(4):
            a,b,z=[x.utils.get_column_letter(c+d) for c in [10,14,18]]
            s[f'{z}{60+i}']=f'={a}{60+i}-{b}{60+i}'
    s['B54']='Grand Total'; s['A90']='Grand Total'
    for c in range(10,22):
        l=x.utils.get_column_letter(c); s.cell(72,c).value=f'=SUM({l}60:{l}71)'
    for c in range(18,22): s.cell(59,c).value=None
    s['P19']='=SUM(P7:P18)'; s['H37']='=SUM(H25:H36)'
    notes=wb.create_sheet('Import Notes'); notes.column_dimensions['A'].width=115
    for line in ['GSTR-3B Offset Summary import', 'Source: '+name, 'Matched months by name (April–March). Source blanks remain blank in imported input cells.', 'Outward total = Outward (NRC) + Non-Taxable; Outward (RC) is excluded as indicated by the REECO heading.', 'As per Books is not imported. Enter books figures before interpreting differences.', 'Totals and differences may show zero or treat blanks as zero; this is not proof that the underlying amounts were reported.', 'Imported: outward values and liability, reverse charge liability, ITC eligible (NRC/RC), reversed, reclaimed and ineligible.']:
        notes.append([line]); notes.cell(notes.max_row,1).data_type='s'
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

if __name__=='__main__':
    import hashlib
    import hmac
    import streamlit as st
    st.title('GST REECO Sheet Automation')
    try:
        password = st.secrets['APP_PASSWORD']
    except (FileNotFoundError, KeyError):
        password = None
    if not isinstance(password, str) or len(password.strip()) < 16 or password == 'REPLACE_WITH_YOUR_PASSWORD':
        st.error('The app is locked. Set APP_PASSWORD in Streamlit Secrets to a unique password of at least 16 characters.')
        st.stop()
    token = hashlib.sha256(password.encode()).digest()
    def sign_in():
        supplied = st.session_state.pop('login_password', '')
        if hmac.compare_digest(supplied.encode(), password.encode()):
            st.session_state['login_token'] = token
            st.session_state.pop('login_failed', None)
        else:
            st.session_state['login_failed'] = True
    if st.session_state.get('login_token') != token:
        with st.form('sign_in'):
            st.text_input('App password', type='password', key='login_password')
            st.form_submit_button('Sign in', on_click=sign_in)
        if st.session_state.get('login_failed'):
            st.error('Incorrect password.')
        st.stop()
    if st.sidebar.button('Sign out'):
        st.session_state.clear()
        st.rerun()
    st.write('Upload a Consolidated 3B Offset Summary to fill the return figures in your REECO template.')
    uploaded=st.file_uploader('3B Offset Summary',type=['xls','xlsx'])
    if uploaded:
        try:
            result=fill(uploaded.getvalue(),uploaded.name)
            st.success('REECO sheet ready. Enter your books figures before reviewing differences.')
            st.download_button('Download filled REECO sheet',result,'REECO Filled.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception as e:
            st.error(f'Could not import this workbook: {e}')
