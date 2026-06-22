import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import os
from datetime import datetime

# ===== 設定 =====
SPREADSHEET_ID = '1h06FfSGadEqViz77rReSlbIs_QIryOE1JloWqjR2GCU'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

EXPENSE_KW = ['旅費交通費','支払手数料','交際費','通信費','消耗品費','リース料','研修費','会議費',
              '地代家賃','外注費','広告宣伝費','水道光熱費','修繕費','保険料','新聞図書費',
              '福利厚生費','雑費','給与手当','給料手当','役員報酬','法定福利費','減価償却費',
              '租税公課','システム利用料','採用費','業務委託費','接待交際費','支払報酬料',
              '支払報酬','諸会費','寄付金','車両費','賞与','退職給与']
SALES_KW = ['売上高']
COGS_KW = ['仕入高', '売上原価']

# ===== Google Sheets接続 =====
@st.cache_resource
def get_gsheet_client():
    # Streamlit Cloud上はSecretsから、ローカルはcredentials.jsonから読込
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES
        )
    except:
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet():
    client = get_gsheet_client()
    return client.open_by_key(SPREADSHEET_ID)

# ===== データ保存・読込 =====
def save_to_sheet(sh, sheet_name, df):
    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=5000, cols=50)
    ws.update([df.columns.tolist()] + df.values.tolist())

def load_from_sheet(sh, sheet_name):
    try:
        ws = sh.worksheet(sheet_name)
        data = ws.get_all_records()
        return pd.DataFrame(data) if data else None
    except:
        return None

# ===== 仕訳帳CSVパース =====
def parse_jn_csv(uploaded_file):
    try:
        content = uploaded_file.read()
        for enc in ('cp932', 'shift-jis', 'utf-8-sig', 'utf-8'):
            try:
                text = content.decode(enc)
                break
            except:
                continue
        else:
            return None, '文字コードを判定できませんでした'

        lines = [l for l in text.split('\n') if l.strip()]
        if len(lines) < 2:
            return None, 'データが少なすぎます'

        import csv, io
        reader = csv.reader(io.StringIO(lines[0]))
        header = next(reader)

        def gi(name):
            try: return header.index(name)
            except: return -1

        iDate=gi('取引日'); iDkAcc=gi('借方勘定科目'); iDkAmt=gi('借方金額')
        iDkDept=gi('借方部門'); iCrAcc=gi('貸方勘定科目'); iCrAmt=gi('貸方金額')
        iCrDept=gi('貸方部門'); iDkPartner=gi('借方取引先名'); iNote=gi('取引内容')

        if iDkAcc < 0 or iDkAmt < 0:
            return None, '仕訳帳CSVの列が見つかりません'

        rows = []
        for line in lines[1:]:
            r = next(csv.reader(io.StringIO(line)))
            if not r: continue
            date_raw = r[iDate].strip() if iDate >= 0 and iDate < len(r) else ''
            if not date_raw or date_raw == 'NaN': continue
            month = date_raw[:7]

            dk_acc = r[iDkAcc].strip() if iDkAcc < len(r) else ''
            dk_amt = float(r[iDkAmt].replace(',','').replace('¥','')) if iDkAmt < len(r) and r[iDkAmt].strip() else 0
            dk_dept = r[iDkDept].strip() if iDkDept >= 0 and iDkDept < len(r) else ''
            dk_dept = '' if dk_dept == 'NaN' else dk_dept
            cr_acc = r[iCrAcc].strip() if iCrAcc >= 0 and iCrAcc < len(r) else ''
            cr_amt = float(r[iCrAmt].replace(',','').replace('¥','')) if iCrAmt >= 0 and iCrAmt < len(r) and r[iCrAmt].strip() else 0
            cr_dept = r[iCrDept].strip() if iCrDept >= 0 and iCrDept < len(r) else ''
            cr_dept = '' if cr_dept == 'NaN' else cr_dept
            partner = r[iDkPartner].strip() if iDkPartner >= 0 and iDkPartner < len(r) else ''
            note = r[iNote].strip()[:40] if iNote >= 0 and iNote < len(r) else ''

            # 売上
            if any(k in cr_acc for k in SALES_KW) and cr_amt > 0:
                rows.append({'date': date_raw, 'month': month, 'type': '売上高',
                             'account': cr_acc, 'dept': cr_dept, 'amount': cr_amt,
                             'partner': partner, 'note': note})
            # 原価
            if any(k in dk_acc for k in COGS_KW) and dk_amt > 0:
                rows.append({'date': date_raw, 'month': month, 'type': '仕入高',
                             'account': dk_acc, 'dept': dk_dept, 'amount': dk_amt,
                             'partner': partner, 'note': note})
            # 費用
            if any(k in dk_acc for k in EXPENSE_KW) and dk_amt > 0:
                rows.append({'date': date_raw, 'month': month, 'type': '費用',
                             'account': dk_acc, 'dept': dk_dept, 'amount': dk_amt,
                             'partner': partner, 'note': note})

        return pd.DataFrame(rows), None
    except Exception as e:
        return None, str(e)

# ===== 集計 =====
def aggregate(df, dept=None, month=None):
    d = df.copy()
    if dept and dept != '全体':
        d = d[d['dept'] == dept]
    if month and month != '累計':
        d = d[d['month'] == month]
    sales = d[d['type']=='売上高']['amount'].sum()
    cogs  = d[d['type']=='仕入高']['amount'].sum()
    exp   = d[d['type']=='費用']['amount'].sum()
    gross = sales - cogs
    op    = gross - exp
    return {'sales': sales, 'cogs': cogs, 'expense': exp, 'gross': gross, 'op': op}

def fmt(n):
    if n == 0: return '¥0'
    sign = '-' if n < 0 else ''
    return f"{sign}¥{abs(int(n)):,}"

def pct(a, b):
    return f"{a/b*100:.1f}%" if b else '-'

# ===== UI =====
st.set_page_config(page_title='業績分析表', layout='wide', page_icon='📈')
st.title('📈 業績分析表')

# サイドバー：データ管理
with st.sidebar:
    st.header('📂 データ管理')

    # 仕訳帳CSV取込
    st.subheader('実績データ（仕訳帳CSV）')
    jn_file = st.file_uploader('freee仕訳帳CSV（新形式）', type='csv', key='jn')
    if jn_file:
        with st.spinner('取込中...'):
            df_new, err = parse_jn_csv(jn_file)
            if err:
                st.error(err)
            else:
                try:
                    sh = get_spreadsheet()
                    save_to_sheet(sh, '実績データ', df_new)
                    st.success(f'✅ {len(df_new)}件を保存しました')
                    st.cache_data.clear()
                except Exception as e:
                    st.error(f'保存エラー: {e}')

    st.divider()
    st.caption('データはGoogle Sheetsに自動保存されます')

# ===== メインデータ読込 =====
@st.cache_data(ttl=60)
def load_data():
    try:
        sh = get_spreadsheet()
        df = load_from_sheet(sh, '実績データ')
        if df is not None and not df.empty:
            df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0)
            return df
    except:
        pass
    return None

df = load_data()

if df is None or df.empty:
    st.info('👈 サイドバーから仕訳帳CSVを取り込んでください')
    st.stop()

# ===== フィルター =====
months = sorted(df['month'].unique().tolist())
depts  = sorted(df[df['dept'] != '']['dept'].unique().tolist())

col1, col2 = st.columns([3, 1])
with col1:
    month_options = months + ['累計']
    sel_month = st.selectbox('月', month_options, index=len(month_options)-2)
with col2:
    dept_options = ['全体'] + depts
    sel_dept = st.selectbox('部門', dept_options)

st.divider()

# ===== サマリーカード =====
agg = aggregate(df, sel_dept, sel_month)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric('売上高', fmt(agg['sales']))
with c2:
    gross_rate = f"粗利率 {pct(agg['gross'], agg['sales'])}"
    st.metric('売上総利益', fmt(agg['gross']), gross_rate)
with c3:
    st.metric('費用合計', fmt(agg['expense']))
with c4:
    op_rate = f"利益率 {pct(agg['op'], agg['sales'])}"
    delta_color = 'normal' if agg['op'] >= 0 else 'inverse'
    st.metric('営業利益', fmt(agg['op']), op_rate, delta_color=delta_color)

st.divider()

# ===== タブ =====
tab1, tab2, tab3 = st.tabs(['📊 費用内訳', '📅 月次詳細', '📈 推移グラフ'])

with tab1:
    d = df.copy()
    if sel_dept != '全体': d = d[d['dept']==sel_dept]
    if sel_month != '累計': d = d[d['month']==sel_month]
    exp_df = d[d['type']=='費用'].groupby('account')['amount'].sum().reset_index()
    exp_df.columns = ['勘定科目', '金額']
    exp_df = exp_df[exp_df['金額']>0].sort_values('金額', ascending=False)
    exp_df['構成比'] = exp_df['金額'].apply(lambda x: pct(x, agg['expense']))
    exp_df['金額'] = exp_df['金額'].apply(fmt)
    st.dataframe(exp_df, use_container_width=True, hide_index=True)

with tab2:
    d = df.copy()
    if sel_dept != '全体': d = d[d['dept']==sel_dept]
    pivot = d.groupby(['month','type','account'])['amount'].sum().reset_index()

    rows = []
    for m in months:
        md = pivot[pivot['month']==m]
        s = md[md['type']=='売上高']['amount'].sum()
        c = md[md['type']=='仕入高']['amount'].sum()
        e = md[md['type']=='費用']['amount'].sum()
        rows.append({'月': m.replace('-','年',1).replace('-','月')+'月',
                     '売上高': int(s), '売上総利益': int(s-c),
                     '費用合計': int(e), '営業利益': int(s-c-e)})
    detail_df = pd.DataFrame(rows)

    # 累計行追加
    total_row = {'月': '累計',
                 '売上高': detail_df['売上高'].sum(),
                 '売上総利益': detail_df['売上総利益'].sum(),
                 '費用合計': detail_df['費用合計'].sum(),
                 '営業利益': detail_df['営業利益'].sum()}
    detail_df = pd.concat([detail_df, pd.DataFrame([total_row])], ignore_index=True)
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

with tab3:
    d = df.copy()
    if sel_dept != '全体': d = d[d['dept']==sel_dept]
    chart_data = []
    for m in months:
        md = d[d['month']==m]
        s = md[md['type']=='売上高']['amount'].sum()
        c = md[md['type']=='仕入高']['amount'].sum()
        e = md[md['type']=='費用']['amount'].sum()
        label = m[5:]+'月'
        chart_data.append({'月': label, '売上高': s, '売上総利益': s-c, '営業利益': s-c-e})
    chart_df = pd.DataFrame(chart_data).set_index('月')
    st.bar_chart(chart_df[['売上高', '売上総利益', '営業利益']])

# ===== 仕訳明細 =====
st.divider()
st.subheader('🔍 仕訳明細')
col_a, col_b = st.columns(2)
with col_a:
    acc_options = sorted(df['account'].unique().tolist())
    sel_acc = st.selectbox('科目', acc_options)
with col_b:
    sel_m2 = st.selectbox('月', ['全期間'] + months, key='detail_month')

d2 = df[df['account']==sel_acc].copy()
if sel_dept != '全体': d2 = d2[d2['dept']==sel_dept]
if sel_m2 != '全期間': d2 = d2[d2['month']==sel_m2]
d2 = d2[['date','account','dept','partner','note','amount']].sort_values('date')
d2.columns = ['日付','科目','部門','取引先','摘要','金額']
d2['金額'] = d2['金額'].apply(fmt)
st.dataframe(d2, use_container_width=True, hide_index=True)
st.caption(f'{len(d2)}件')
