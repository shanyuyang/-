import sys
import streamlit as st
import pandas as pd
import numpy as np
import io
import calendar
import re
from datetime import datetime

# —— 页面配置 ——
st.set_page_config(page_title='出荷分析App', layout='wide')
st.title("📊 出荷分析")

# —— 语言切换 ——
st.sidebar.markdown("### 🌐 语言 / Language")
lang = st.sidebar.selectbox("📢 语言 / 言語", ["中文", "日本語"], index=0)
def t(zh, jp): return jp if lang == "日本語" else zh

# —— 上传 & 缓存原始数据（支持清洗页传入） ——
if 'df' not in st.session_state:
    # ✅ 若清洗页传入 cleaned_df，则使用它作为分析输入
    if 'cleaned_df' in st.session_state:
        st.session_state.df = st.session_state.cleaned_df
    else:
        upload = st.sidebar.file_uploader(
            t('上传 原始 数据 Excel/CSV (含“原始データ”sheet)', '原始データ をアップロード'),
            type=['xlsx', 'csv']
        )
        if not upload:
            st.info(t('请上传文件后再运行', 'ファイルをアップロードしてください'))
            st.stop()
        try:
            df = (pd.read_csv(upload, encoding='utf-8-sig', dtype=str)
                  if upload.name.lower().endswith('.csv')
                  else pd.read_excel(upload, sheet_name=0, engine='openpyxl', dtype=str))
        except Exception as e:
            st.error(t(f'读取失败: {e}', '読み込み失敗: {e}'))
            st.stop()
        st.session_state.df = df

df = st.session_state.df.copy()

# —— 清洗列名 ——
df.columns = [c.strip().lower() for c in df.columns]

# —— 定义辅助解析函数 ——
def normalize_ym(x, default_year):
    s = str(x).strip()
    # 形如 '2025年1月'
    m = re.match(r"(\d{4})\D*(\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    # 形如 '1月'
    m2 = re.match(r"(\d{1,2})月", s)
    if m2:
        return f"{default_year:04d}-{int(m2.group(1)):02d}"
    # 已是 'YYYY-MM'
    parts = s.split('-')
    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}"
    raise ValueError(f"无法解析年月格式：{x}")

def parse_year_month(period, default_year):
    # 从标准化后的 'YYYY-MM' 或 '1月' 提取 year, month
    nums = re.findall(r"\d+", str(period))
    if len(nums) == 2:
        return int(nums[0]), int(nums[1])
    elif len(nums) == 1:
        return default_year, int(nums[0])
    else:
        raise ValueError(f"解析年月失败: {period}")

def expected_union(days, d1, d2, d3):
    # 计算三仓库断货天数并集期望
    p1, p2, p3 = d1/days, d2/days, d3/days
    return days * (1 - (1-p1)*(1-p2)*(1-p3))

# —— 规范化“年月”列 ——
first_val = str(df['年月'].dropna().iloc[0])
match_year = re.search(r"(\d{4})", first_val)
default_year = int(match_year.group(1)) if match_year else datetime.now().year
df['年月'] = df['年月'].apply(lambda x: normalize_ym(x, default_year))
df['年月'] = pd.to_datetime(df['年月'], format='%Y-%m', errors='coerce').dt.to_period('M')

# —— 核心列自动映射（识别不到再选） —— 
col_patterns = {
    '年月':    ['年月','date','month','ym'],
    '工場名':  ['工場名','工厂','factory'],
    '贸易公司':['贸易公司','商社名','company'],
    '商品编码':['商品编码','商品コード','sku','code'],
    '欠品次数':['欠品次数','欠品回数','stockout','回数','欠品理由_合計'],
    '商品名称':['商品名称','产品名称','商品名','name'],
    '出货次数':['出货次数','出貨回数','ship_count','shipping'],
    '产品状态':['产品状态','商品状态','status']
}


for target, pats in col_patterns.items():
    # 在所有列名中找包含任一模式的列
    matches = [c for c in df.columns if any(p.lower() in c for p in pats)]
    if len(matches) == 1:
        # 唯一命中，直接重命名
        df.rename(columns={matches[0]: target}, inplace=True)
    else:
        # 多选或未选中，弹框让用户手动选
        sel = st.sidebar.selectbox(
            t(f'请选择“{target}”列', f'「{target}」列を選択'),
            df.columns.tolist(), key=target
        )
        df.rename(columns={sel: target}, inplace=True)

if '产品状态' in df.columns:
    # 先把含“新”的全标为 1，含“廃番"/"废番"的标为 2，其余留原值
    def map_status(x):
        if isinstance(x, str):
            if '新' in x: return 1
            if '廃番' in x or '废番' in x: return 2
        try:
            v = int(x)
            return v if v in (0,1,2) else 0
        except:
            return 0
    
    df['产品状态'] = df['产品状态'].apply(map_status)
else:
    # 源表里没有这列，就一律当“在产产品”
    df['产品状态'] = 0

# 在这里定义必需列列表，用于后续校验
required_cols = list(col_patterns.keys())   

# 校验映射完成
missing = [col for col in required_cols if col not in df.columns]
if missing:
    st.warning(t(f'缺少核心列: {missing}', f'不足している主要列: {missing}'))
    st.stop()

# —— 仓库断货天数自动映射（识别不到再选） —— 
for wh in ['0001','1001','2001']:
    target_col = f'{wh}仓库欠品日数'
    # 必须同时含有仓库编号和欠品日数关键字
    pats = [wh, '欠品', '日数']
    matches = [c for c in df.columns if all(p in c for p in pats)]
    if len(matches) == 1:
        df.rename(columns={matches[0]: target_col}, inplace=True)
    else:
        sel = st.sidebar.selectbox(
            t(f'请选择 {wh} 仓库断货天数 列', f'{wh}倉庫欠品日数 列を選択'),
            df.columns.tolist(), key=target_col
        )
        df.rename(columns={sel: target_col}, inplace=True)


# —— 类型转换 —— ——
ignore_cols = ['年月','工場名','贸易公司','商品编码','商品名称','仓库编号']
num_cols = [c for c in df.columns if c not in ignore_cols]
for c in num_cols:
    df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)





# —— 预计/实际 列映射 ——
st.sidebar.markdown("### 📊 计算列映射")
pre_key = st.sidebar.selectbox(t('请选择预计列','予定数列を選択'), num_cols, index=0)
ship_key = st.sidebar.selectbox(t('请选择实际列','実績数列を選択'), num_cols, index=0)

# —— 异常校验 —— 
invalid = df[df[ship_key] > df[pre_key]]
if not invalid.empty:
    # ① 完整闭合的 warning
    st.warning(
        t(
            '存在实际出货超过预计的数据',
            '実際出荷が予定を超えたデータがあります'
        )
    )
    # ② 完整闭合的 expander
    with st.expander(
        t('🚩 超出预计记录', '🚩 予定超過レコード'),
        expanded=False
    ):
        # ③ datafram e跨行写时，也要包在括号里
        st.dataframe(
            invalid[
                ['贸易公司', '工場名', '商品编码', pre_key, ship_key]
            ],
            use_container_width=True
        )

# 原始“欠品日数”列已删除，不再进行相关校验
if (df['欠品次数'] < 0).any():
    st.error(t('存在负值欠品次数','欠品次数に負の値があります'))

# —— 筛选 ——
months = sorted(df['年月'].unique())
with st.sidebar.expander(t('🔎 筛选','🔎 フィルター'), True):
    sel_trade = st.selectbox(t('贸易公司','貿易会社'), ['全部']+list(df['贸易公司'].unique()))
    if sel_trade!='全部': df = df[df['贸易公司']==sel_trade]
    sel_fac = st.selectbox(t('工厂','工場'), ['全部']+list(df['工場名'].unique()))
    if sel_fac!='全部': df = df[df['工場名']==sel_fac]
    sel_sku = st.selectbox(t('SKU','SKU'), ['全部']+list(df['商品编码'].unique()))
    if sel_sku!='全部': df = df[df['商品编码']==sel_sku]
    sel_months = st.multiselect(t('年月','年月'), months, default=months)
    if sel_months: df = df[df['年月'].isin(sel_months)]


# —— SKU 汇总与达成率计算 ——
def calc_monthly_rate(df_m, by_cols):
    g = df_m.groupby(by_cols).agg(pre=(pre_key,'sum'), sh=(ship_key,'sum')).reset_index()
    g['rate'] = np.where(g['pre']>0, (g['sh']/g['pre']).round(4), np.nan)
    return g

sku_grp = df.groupby(['贸易公司','工場名','商品编码']).agg(
    total_pre=(pre_key,'sum'),
    total_sh=(ship_key,'sum'),
    欠品次数合计=('欠品次数','sum')
).reset_index()
sku_grp['产品名称'] = df.groupby(['贸易公司','工場名','商品编码'])['商品名称'].first().values
rate_cols, weight_cols = [], []
for m in sel_months:
    tmp = calc_monthly_rate(df[df['年月']==m], ['贸易公司','工場名','商品编码'])
    tmp = tmp.rename(columns={'rate':f'{m}达成率','pre':f'{m}预订量'})
    sku_grp = sku_grp.merge(tmp[['贸易公司','工場名','商品编码',f'{m}达成率',f'{m}预订量']],
                              on=['贸易公司','工場名','商品编码'], how='left')
    rate_cols.append(f'{m}达成率'); weight_cols.append(f'{m}预订量')

# —— 断货日数比率修正 ——
# 计算各月真实天数
parsed = [parse_year_month(m, default_year) for m in sel_months]
days_in_month = {m: calendar.monthrange(y, mo)[1] for m, (y, mo) in zip(sel_months, parsed)}
# 提取月度各仓库欠品日数列
wh_cols = ['0001仓库欠品日数','1001仓库欠品日数','2001仓库欠品日数']
month_wh = df[df['年月'].isin(sel_months)][
    ['贸易公司','工場名','商品编码','年月'] + wh_cols
]
# 期望并集天数
month_wh['union_days'] = month_wh.apply(
    lambda r: expected_union(
        days_in_month[r['年月']],
        r['0001仓库欠品日数'],
        r['1001仓库欠品日数'],
        r['2001仓库欠品日数']
    ), axis=1
)
# 汇总至 SKU 维度
stockout_sum = (
    month_wh.groupby(['贸易公司','工場名','商品编码'], as_index=False)['union_days']
    .sum()
    .rename(columns={'union_days':'sku_stockout_days'})
)
sku_grp = sku_grp.merge(stockout_sum, on=['贸易公司','工場名','商品编码'], how='left').fillna({'sku_stockout_days':0})
# 计算比率
total_days = sum(days_in_month.values())
sku_grp['断货日数比率'] = (sku_grp['sku_stockout_days'] / total_days).round(4)# ✅ 缺货频率（月度）逻辑

sku_grp['实际断货天数（天）'] = sku_grp['sku_stockout_days']

def count_stockout_months(r):
    return sum([(r[f] < 1 if not pd.isna(r[f]) else False) for f in rate_cols])

sku_grp['缺货频率（月度）'] = (sku_grp.apply(count_stockout_months, axis=1) / len(rate_cols)).round(4)

# ✅ 缺货频率（月度）逻辑

def count_stockout_months(r):
    return sum([(r[f] < 1 if not pd.isna(r[f]) else False) for f in rate_cols])

sku_grp['缺货频率（月度）'] = (sku_grp.apply(count_stockout_months, axis=1) / len(rate_cols)).round(4)

# —— 先按月计算“月度缺货频率（次数）”，再求月度平均 —— 
# 1）按年月 + SKU 维度聚合：算月度欠品次数 & 月度出货次数
monthly_evt = (
    df
    .groupby(['年月','贸易公司','工場名','商品编码'])
    .agg(
        month_stockout=('欠品次数','sum'),
        month_ship   =('出货次数','sum')
    )
    .reset_index()
)
# 2）计算“月度频率（次数）”：只有当月出货>0 时才算，否则 NaN
monthly_evt['月度频率（次数）'] = np.where(
    monthly_evt['month_ship'] > 0,
    (monthly_evt['month_stockout'] / monthly_evt['month_ship']).round(4),
    np.nan
)
# 3）对每个 SKU 求算术平均
avg_evt = (
    monthly_evt
    .groupby(['贸易公司','工場名','商品编码'])['月度频率（次数）']
    .mean()
    .reset_index(name='平均月度缺货频率（次数）')
)
# 4）合并回 sku_grp
sku_grp = sku_grp.merge(
    avg_evt,
    on=['贸易公司','工場名','商品编码'],
    how='left'
)

# —— 初始化空字段避免 KeyError ——

sku_grp['产品状态'] = ''
if '产品状态' in df.columns:
    sku_grp['产品状态'] = df.groupby(['贸易公司','工場名','商品编码'])['产品状态'].first().values
else:
    sku_grp['产品状态'] = '0'  # 若源数据中无此列，则设为空

if 'total_pre' in sku_grp.columns and 'total_sh' in sku_grp.columns:
    sku_grp['总达成率'] = np.where(
        sku_grp['total_pre'] > 0,
        (sku_grp['total_sh'] / sku_grp['total_pre']).round(4),
        np.nan
    )
else:
    sku_grp['总达成率'] = np.nan  # 防止 total_pre 缺失导致错误



# ✅ 其他指标字段计算
def safe_weighted_avg(row):
    values = []
    weights = []
    for r, w in zip(rate_cols, weight_cols):
        val = row[r]
        weight = row[w]
        if pd.notna(val) and pd.notna(weight):
            values.append(val * weight)
            weights.append(weight)
    if sum(weights) > 0:
        return sum(values) / sum(weights)
    else:
        return np.nan

# —— 1）先把月度预订量和达成率中的 NaN 当作 0 —— 
sku_grp[weight_cols] = sku_grp[weight_cols].fillna(0)
sku_grp[rate_cols]   = sku_grp[rate_cols].fillna(0)

# —— 直接用 app.py 的加权平均公式 —— 
num = sum(sku_grp[w] * sku_grp[r] for w, r in zip(weight_cols, rate_cols))
den = sku_grp[weight_cols].sum(axis=1)
sku_grp['实际平均达成率（加权）'] = np.where(den > 0, (num/den).round(4), np.nan)

# ✅ 同时计算 平均达成率（简单平均）
sku_grp['平均达成率'] = sku_grp[rate_cols].mean(axis=1).round(4)


# ✅ 字段中日统一命名（重命名）
rename_dict = {
    '产品状态': '产品状态' if lang == '中文' else '商品状態',
    '贸易公司': '贸易公司' if lang == '中文' else '貿易会社',
    '工場名': '工厂名' if lang == '中文' else '工場名',
    '商品编码': '商品编码' if lang == '中文' else '商品コード',
    '产品名称': '产品名称' if lang == '中文' else '商品名',
    '总达成率': '总达成率' if lang == '中文' else '総達成率',
    '平均达成率': '平均达成率' if lang == '中文' else '平均達成率',
    '实际平均达成率（加权）': '实际平均达成率（加权）' if lang == '中文' else '実質平均達成率（加重）',
    '断货日数比率': '断货日数比率' if lang == '中文' else '欠品日数比率',
    '平均月度缺货频率（次数）': '平均月度缺货频率（次数）' if lang=='中文' else '平均月次欠品頻度（回数）',
    '缺货频率（月度）': '缺货频率（月度）' if lang == '中文' else '欠品頻度（月次）',
    '实际断货天数（天）': '实际断货天数（天）' if lang == '中文' else '実際欠品日数（日）',
    '出货次数': '出货次数' if lang == '中文' else '出荷回数',
    '欠品次数合计': '欠品次数合计' if lang == '中文' else '欠品回数合計'
}
sku_grp.rename(columns=rename_dict, inplace=True)

# ✅ 字段顺序调整（缺货频率（月度）放在第11位）
display_sku = [
    rename_dict['产品状态'], rename_dict['贸易公司'], rename_dict['工場名'], rename_dict['商品编码'], rename_dict['产品名称'],
    rename_dict['总达成率'], rename_dict['平均达成率'], rename_dict['实际平均达成率（加权）'],
    rename_dict['断货日数比率'], rename_dict['平均月度缺货频率（次数）'], rename_dict['缺货频率（月度）'], rename_dict['欠品次数合计']
] + rate_cols + [
    rename_dict['实际断货天数（天）']
]


# ✅ 展示表格
st.subheader(t('▶ SKU 分析（按综合风险排序）','▶ SKU 分析（総合リスク順）'))

# 不再筛状态，只对整个 sku_grp 排序
top_n = st.number_input(
    '显示最差前 N 个 SKU',
    min_value=1,
    max_value=len(sku_grp),
    value=min(20, len(sku_grp)),
    step=1
)

sku_sorted = sku_grp.sort_values(
    by=[
        '产品状态',                    # 0 在前，1、2 在后
        '实际平均达成率（加权）',      # 状态相同时，再看这三项
        '断货日数比率',
        '平均月度缺货频率（次数）'
    ],
    ascending=[True, True, False, False]
).reset_index(drop=True)

sku_sorted.insert(0, '排名', sku_sorted.index + 1)
to_show = sku_sorted.head(top_n)
st.dataframe(to_show[['排名'] + display_sku], use_container_width=True)


# —— 工厂 分析 —— 
fac_grp = df.groupby(['贸易公司','工場名']).agg(
    total_pre=('{}'.format(pre_key),'sum'),
    total_sh =('{}'.format(ship_key),'sum'),
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

fac_grp.drop(columns=weight_cols, inplace=True)

st.markdown('---')
st.subheader(t('▶ 工厂 分析','▶ 工場 分析'))
display_fac=['贸易公司','工場名']+rate_cols+[ 
    '总达成率','平均达成率','加权平均达成率',
    'SKU数','欠品次数合計'
]
st.dataframe(fac_grp[display_fac],use_container_width=True)

# —— 贸易公司 分析 —— 
trade_grp = df.groupby('贸易公司').agg(
    total_pre=('{}'.format(pre_key),'sum'),
    total_sh =('{}'.format(ship_key),'sum'),
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
trade_grp.drop(columns=weight_cols, inplace=True)


st.markdown('---')
st.subheader(t('▶ 贸易公司 分析','▶ 貿易会社 分析'))
display_trade=['贸易公司']+rate_cols+[ 
    '总达成率','平均达成率','加权平均达成率',
    'SKU数','欠品次数合計'
]
st.dataframe(trade_grp[display_trade],use_container_width=True)

# —— 下载按钮 & Excel 报告 —— 
st.sidebar.markdown("### 📥 下载")

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine='openpyxl') as writer:
    # 只导出页面上“最差前 N 个 SKU”以及页面展示的列
    cols = ['排名'] + display_sku
    to_show[cols].to_excel(writer, sheet_name='SKU分析', index=False)
    fac_grp[display_fac].to_excel(writer, sheet_name='工厂分析', index=False)
    trade_grp[display_trade].to_excel(writer, sheet_name='公司分析', index=False)
    df.to_excel(writer, sheet_name='原数据', index=False)


st.download_button(
    t('下载 全量 Excel', 'Excelダウンロード'),
    buf,
    'report.xlsx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
