import sys
import streamlit as st
import pandas as pd
import numpy as np
import io

# —— 页面配置 —— 
st.set_page_config(page_title='出荷分析App', layout='wide')

# —— 语言切换 —— 
st.sidebar.markdown("### 🌐 语言 / Language")
lang = st.sidebar.selectbox("📢 语言 / 言語", ["中文", "日本語"], index=0)
def t(zh, jp): return jp if lang == "日本語" else zh

# —— 上传 & 缓存原始数据到 session_state —— 
if 'df' not in st.session_state:
    upload = st.sidebar.file_uploader(
        t('上传 原始 数据 Excel/CSV (含“原始データ”sheet 或 首行为列头)', 
          '原始データ Excel/CSV をアップロード'),
        type=['xlsx', 'csv']
    )
    if not upload:
        st.info(t('请上传文件后再运行', 'ファイルをアップロードしてください'))
        st.stop()
    try:
        if upload.name.lower().endswith('.csv'):
            df = pd.read_csv(upload, encoding='utf-8-sig', dtype=str)
        else:
            df = pd.read_excel(upload, sheet_name='原始データ', 
                               engine='openpyxl', dtype=str)
    except Exception as e:
        st.error(t(f'读取失败: {e}', '読み込み失敗: {e}'))
        st.stop()
    st.session_state.df = df

# —— 读取缓存的源数据 —— 
df = st.session_state.df.copy()

# —— 清洗列名 —— 
df.columns = [c.strip().lower() for c in df.columns]  # 去除空格和特殊字符

# —— 确保列映射正确 —— 
fields = [
    ('年月', '年月', '年月'),
    ('工場名', '工厂', '工場名'),
    ('贸易公司', '贸易公司', '貿易会社'),
    ('商品编码', '商品编码', '商品コード'),
    ('Pre-Order CTN数', '预订箱数', 'Pre-Order CTN数'),
    ('Pre-Order PCS数', '预订PCS数', 'Pre-Order PCS数'),
    ('出荷可能 CTN数', '可出箱数', '出荷可能 CTN数'),
    ('出荷可能 PCS数', '可出PCS数', '出荷可能 PCS数'),
    ('正式オーダー CTN数', '正式订单箱数', '正式オーダー CTN数'),
    ('正式オーダー PCS数', '正式订单PCS数', '正式オーダー PCS数'),
    ('欠品日数', '欠品日数', '欠品日数'),
    ('欠品次数', '欠品次数', '欠品次数'),
    ('欠品理由', '缺货原因', '欠品理由'),
    ('商品名称', '产品名称', '商品名称')  # 确保产品名称映射正确
]

# 确保每一列都能正确映射
for col, zh, jp in fields:
    if col not in df.columns:
        sel = st.sidebar.selectbox(
            t(f'请选择“{zh}”列', f'「{jp}」列を選択'),
            df.columns.tolist(), key=col)
        df.rename(columns={sel: col}, inplace=True)

missing = [col for col, _zh, _jp in fields if col not in df.columns]
if missing:
    st.warning(t(f'缺少列: {missing}', f'不足している列: {missing}'))
    st.stop()

# —— 类型转换 —— 
df['年月'] = df['年月'].astype(str).str.strip()
df = df[df['年月'] != '']
num_cols = [c for c in df.columns if c not in ['年月', '工場名', '贸易公司', '商品编码', '欠品理由', '商品名称']]
for c in num_cols:
    df[c] = pd.to_numeric(df[c].str.replace(',', ''), errors='coerce').fillna(0)

# —— 预计/实际 列映射 —— 
st.sidebar.markdown("### 📊 计算列映射")
pre_key = st.sidebar.selectbox(
    t('请选择预计列', '予定数列を選択'), num_cols,
    index=num_cols.index('Pre-Order PCS数') if 'Pre-Order PCS数' in num_cols else 0)
ship_key = st.sidebar.selectbox(
    t('请选择实际列', '実績数列を選択'), num_cols,
    index=num_cols.index('正式オーダー PCS数') if '正式オーダー PCS数' in num_cols else 0)

# —— 异常校验 —— 
invalid = df[df[ship_key] > df[pre_key]]
if not invalid.empty:
    st.warning(t(f"存在 {len(invalid)} 条 实际 > 预计", "実績数が予定数超過"))
    st.subheader(t('🚩 超出预计记录', '🚩 予定超過レコード'))
    st.dataframe(invalid[['贸易公司', '工場名', '商品编码', pre_key, ship_key, '欠品理由']],
                 use_container_width=True)

if (df['欠品日数'] < 0).any() or (df['欠品次数'] < 0).any():
    st.error(t('存在负值欠品', '欠品に負の値があります'))

# —— 筛选 —— 
months = sorted(df['年月'].unique())
with st.sidebar.expander(t('🔎 筛选', '🔎 フィルター'), True):
    sel_trade = st.selectbox(t('贸易公司', '貿易会社'), ['全部'] + sorted(df['贸易公司'].unique()))
    if sel_trade != '全部': df = df[df['贸易公司'] == sel_trade]
    sel_fac = st.selectbox(t('工厂', '工場'), ['全部'] + sorted(df['工場名'].unique()))
    if sel_fac != '全部': df = df[df['工場名'] == sel_fac]
    sel_sku = st.selectbox(t('SKU', 'SKU'), ['全部'] + sorted(df['商品编码'].unique()))
    if sel_sku != '全部': df = df[df['商品编码'] == sel_sku]
    sel_months = st.multiselect(t('年月', '年月'), months, default=months)
    if sel_months: df = df[df['年月'].isin(sel_months)]

# —— 辅助函数 —— 
def calc_monthly_rate(df_m, by_cols):
    g = df_m.groupby(by_cols).agg(pre=(pre_key, 'sum'), sh=(ship_key, 'sum')).reset_index()
    g['rate'] = np.where(g['pre'] > 0, (g['sh'] / g['pre']).round(4), np.nan)
    return g

# —— SKU 分析 —— 
sku_grp = df.groupby(['贸易公司', '工場名', '商品编码']).agg(
    total_pre=('{}'.format(pre_key), 'sum'),
    total_sh=('{}'.format(ship_key), 'sum'),
    欠品日数合计=('欠品日数', 'sum'),
    欠品次数合计=('欠品次数', 'sum')
).reset_index()

# 添加产品名称列
sku_grp['产品名称'] = df.groupby(['贸易公司', '工場名', '商品编码'])['商品名称'].first().reset_index(drop=True)

rate_cols, weight_cols = [], []
for m in sel_months:
    tmp = calc_monthly_rate(df[df['年月'] == m], ['贸易公司', '工場名', '商品编码'])
    tmp = tmp.rename(columns={'rate': f'{m}达成率', 'pre': f'{m}预订量'})
    sku_grp = sku_grp.merge(tmp[['贸易公司', '工場名', '商品编码', f'{m}达成率', f'{m}预订量']],
                            on=['贸易公司', '工場名', '商品编码'], how='left')
    rate_cols.append(f'{m}达成率')
    weight_cols.append(f'{m}预订量')

sku_grp['总达成率'] = np.where(sku_grp['total_pre'] > 0,
                               (sku_grp['total_sh'] / sku_grp['total_pre']).round(4), np.nan)
sku_grp['平均达成率'] = sku_grp[rate_cols].mean(axis=1).round(4)
num = sum(sku_grp[w] * sku_grp[r] for w, r in zip(weight_cols, rate_cols))
den = sku_grp[weight_cols].sum(axis=1)
sku_grp['加权平均达成率'] = np.where(den > 0, (num / den).round(4), np.nan)
sku_grp['达成率标准差'] = sku_grp[rate_cols].std(axis=1).round(4)
sku_grp['达成率变异系数'] = np.where(
    sku_grp['平均达成率'] > 0,
    (sku_grp['达成率标准差'] / sku_grp['平均达成率']).round(4),
    np.nan
)
sku_grp.drop(columns=weight_cols, inplace=True)

# —— SKU 新增衍生指标 —— 
sku_grp['MAD'] = sku_grp[rate_cols].apply(
    lambda r: np.nanmean(np.abs(r - np.nanmean(r))), axis=1).round(4)
total_days = len(sel_months) * 30
sku_grp['断货日数比率'] = (sku_grp['欠品日数合计'] / total_days).round(4)
sku_grp['发运月数'] = sku_grp[rate_cols].gt(0).sum(axis=1)
sku_grp['缺货频率'] = np.where(
    sku_grp['发运月数'] > 0,
    (sku_grp['欠品次数合计'] / sku_grp['发运月数']).round(4),
    np.nan
)
sku_grp.drop(columns=['发运月数'], inplace=True)

# —— 展示 SKU 分析 —— 
display_sku = ['贸易公司', '工場名', '商品编码', '产品名称'] + rate_cols + [
    '总达成率', '平均达成率', '加权平均达成率',
    '达成率标准差', '达成率变异系数',
    'MAD', '断货日数比率', '缺货频率',
    '欠品日数合计', '欠品次数合计'
]

st.subheader(t('▶ SKU 分析', '▶ SKU 分析'))
st.dataframe(sku_grp[display_sku], use_container_width=True)

# —— 工厂 分析 —— 
fac_grp = df.groupby(['贸易公司','工場名']).agg(
    total_pre=('{}'.format(pre_key),'sum'),
    total_sh =('{}'.format(ship_key),'sum'),
    欠品日数合計=('欠品日数','sum'),
    欠品次数合計=('欠品次数','sum'),
    SKU数      =('商品编码','nunique')
).reset_index()
for m in sel_months:
    tmp = calc_monthly_rate(df[df['年月']==m],['贸易公司','工場名'])
    tmp=tmp.rename(columns={'rate':f'{m}达成率','pre':f'{m}预订量'})
    fac_grp=fac_grp.merge(tmp[['贸易公司','工場名',f'{m}达成率',f'{m}预订量']],
                          on=['贸易公司','工場名'],how='left')

fac_grp['总达成率']=np.where(fac_grp['total_pre']>0,
                            (fac_grp['total_sh']/fac_grp['total_pre']).round(4),np.nan)
fac_grp['平均达成率']=fac_grp[rate_cols].mean(axis=1).round(4)
num_f=sum(fac_grp[w]*fac_grp[r] for w,r in zip(weight_cols,rate_cols))
den_f=fac_grp[weight_cols].sum(axis=1)
fac_grp['加权平均达成率']=np.where(den_f>0,(num_f/den_f).round(4),np.nan)
fac_grp['达成率标准差']=fac_grp[rate_cols].std(axis=1).round(4)
fac_grp['达成率变异系数']=np.where(
    fac_grp['平均达成率']>0,
    (fac_grp['达成率标准差']/fac_grp['平均达成率']).round(4),
    np.nan
)
fac_grp.drop(columns=weight_cols, inplace=True)

# —— 工厂 新增 MAD —— 
fac_grp['MAD']=fac_grp[rate_cols].apply(
    lambda r: np.nanmean(np.abs(r-np.nanmean(r))), axis=1).round(4)

st.markdown('---')
st.subheader(t('▶ 工厂 分析','▶ 工場 分析'))
display_fac=['贸易公司','工場名']+rate_cols+[ 
    '总达成率','平均达成率','加权平均达成率',
    '达成率标准差','达成率变异系数','MAD',
    'SKU数','欠品日数合計','欠品次数合計'
]
st.dataframe(fac_grp[display_fac],use_container_width=True)

# —— 贸易公司 分析 —— 
trade_grp = df.groupby('贸易公司').agg(
    total_pre=('{}'.format(pre_key),'sum'),
    total_sh =('{}'.format(ship_key),'sum'),
    欠品日数合計=('欠品日数','sum'),
    欠品次数合計=('欠品次数','sum'),
    SKU数      =('商品编码','nunique')
).reset_index()
for m in sel_months:
    tmp = calc_monthly_rate(df[df['年月']==m],['贸易公司'])
    tmp=tmp.rename(columns={'rate':f'{m}达成率','pre':f'{m}预订量'})
    trade_grp=trade_grp.merge(tmp[['贸易公司',f'{m}达成率',f'{m}预订量']],
                              on=['贸易公司'],how='left')

trade_grp['总达成率']=np.where(trade_grp['total_pre']>0,
                              (trade_grp['total_sh']/trade_grp['total_pre']).round(4),np.nan)
trade_grp['平均达成率']=trade_grp[rate_cols].mean(axis=1).round(4)
num_t=sum(trade_grp[w]*trade_grp[r] for w,r in zip(weight_cols,rate_cols))
den_t=trade_grp[weight_cols].sum(axis=1)
trade_grp['加权平均达成率']=np.where(den_t>0,(num_t/den_t).round(4),np.nan)
trade_grp['达成率标准差']=trade_grp[rate_cols].std(axis=1).round(4)
trade_grp['达成率变异系数']=np.where(
    trade_grp['平均达成率']>0,
    (trade_grp['达成率标准差']/trade_grp['平均达成率']).round(4),
    np.nan
)
trade_grp.drop(columns=weight_cols, inplace=True)

# —— 贸易公司 新增 MAD —— 
trade_grp['MAD']=trade_grp[rate_cols].apply(
    lambda r: np.nanmean(np.abs(r-np.nanmean(r))), axis=1).round(4)

st.markdown('---')
st.subheader(t('▶ 贸易公司 分析','▶ 貿易会社 分析'))
display_trade=['贸易公司']+rate_cols+[ 
    '总达成率','平均达成率','加权平均达成率',
    '达成率标准差','达成率变异系数','MAD',
    'SKU数','欠品日数合計','欠品次数合計'
]
st.dataframe(trade_grp[display_trade],use_container_width=True)

# —— 下载按钮 & Excel 报告 —— 
st.sidebar.markdown("### 📥 下载")
st.download_button(t('下载 SKU CSV','CSVダウンロード'),
                   sku_grp.to_csv(index=False,encoding='utf-8-sig'),'sku.csv')
st.download_button(t('下载 工厂 CSV','CSVダウンロード'),
                   fac_grp.to_csv(index=False,encoding='utf-8-sig'),'factory.csv')
st.download_button(t('下载 公司 CSV','CSVダウンロード'),
                   trade_grp.to_csv(index=False,encoding='utf-8-sig'),'trade.csv')

buf=io.BytesIO()
with pd.ExcelWriter(buf,engine='openpyxl') as writer:
    sku_grp .to_excel(writer,sheet_name='SKU分析', index=False)
    fac_grp .to_excel(writer,sheet_name='工厂分析', index=False)
    trade_grp.to_excel(writer,sheet_name='公司分析', index=False)
    df      .to_excel(writer,sheet_name='原数据', index=False)
buf.seek(0)
st.download_button(t('下载 全量 Excel','Excel下载'),
                   buf,'report.xlsx',
                   'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
