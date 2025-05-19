import io
import pandas as pd
import streamlit as st
import re

st.set_page_config(page_title="📦 欠品 × 出荷 統合ツール", layout="wide")
st.title("📦 欠品日数 × 出荷データ × 商品状態 統合ツール")

file_uploaded = st.file_uploader(" 統合データファイル（複数シート構成）", type=["xlsx"])

if file_uploaded:
    try:
        xls = pd.ExcelFile(file_uploaded)
        df_ketsu_raw = None
        df_ship_raw = None
        df_sales_raw = None
        df_status = pd.DataFrame()

        for sheet_name in xls.sheet_names:
            df_tmp = pd.read_excel(xls, sheet_name=sheet_name)
            columns = df_tmp.columns
            columns_str = columns.astype(str)

            if df_ship_raw is None and '商品コード' in columns_str and any('日付' in str(c) or '出荷' in str(c) for c in columns_str):
                df_ship_raw = df_tmp
                st.success(f"✔️ 出荷データを検出：{sheet_name}")

            elif df_ketsu_raw is None and '商品コード' in columns_str and any(re.match(r'\d{1,2}月(0001|1001|2001)', str(col)) for col in columns_str):
                df_ketsu_raw = df_tmp
                st.success(f"✔️ 欠品データを検出：{sheet_name}")

            elif df_sales_raw is None and '商品コード' in columns_str:
                count_date_cols = sum(bool(re.search(r'20\d{2}[/年\-]\d{1,2}', str(col))) for col in columns_str)
                if count_date_cols >= 3:
                    df_sales_raw = df_tmp
                    st.success(f"✔️ 販売量データを検出：{sheet_name}")

            columns_str = df_tmp.columns.astype(str)
        if '商品コード' in columns_str:
            # 识别所有列名中带有关键字的列
            status_cols = [
                col for col in columns_str
                if any(keyword in col for keyword in ['区分', '状態', '理由', '状态'])
            ]
            # 遍历这些列，逐条提取到 df_status
            for sc in status_cols:
                # 抽取“商品コード” + 该状态列
                tmp = df_tmp[['商品コード', sc]].dropna(subset=[sc])
                # 重命名：把 sc 统一成 '商品状態'
                tmp = tmp.rename(columns={sc: '商品状態'})
                df_status = pd.concat([df_status, tmp], ignore_index=True)

    except Exception as e:
        st.error(f"❌ ファイル読み込み中にエラーが発生しました: {e}")
        st.stop()

    if df_ship_raw is None or df_ketsu_raw is None:
        st.error("❌ 欠品または出荷データが見つかりません。Excelファイルをご確認ください。")
        st.stop()

    months = [f"{i}月" for i in range(1, 13)]
    warehouse_codes = ["0001", "1001", "2001"]
    ketsu_records = []
    for _, row in df_ketsu_raw.iterrows():
        code = row.get("商品コード")
        if pd.isna(code):
            continue
        for month in months:
            vals = []
            for w in warehouse_codes:
                col_name = f"{month}{w}"
                val = pd.to_numeric(row.get(col_name, 0), errors='coerce') or 0
                vals.append(val)
            total = sum(vals)
            ketsu_records.append({
                "商品コード": str(code).strip(),
                "年月": month,
                "0001倉庫欠品日数": vals[0],
                "1001倉庫欠品日数": vals[1],
                "2001倉庫欠品日数": vals[2],
                "三倉庫欠品日数合計": total
            })
    df_ketsu = pd.DataFrame(ketsu_records)

    df_ship = df_ship_raw.copy()
    st.sidebar.title("出荷データ列マッピング")
    date_col = st.sidebar.selectbox("日付列（年月用）", df_ship.columns)
    code_col = st.sidebar.selectbox("商品コード列", df_ship.columns)
    trade_col = st.sidebar.selectbox("貿易会社列", df_ship.columns)

    # 商品名列の自動検出と標準化（修复逻辑，确保后续能用）
    possible_name_cols = ['商品名', '品名', '商品名称']
    for name_col in possible_name_cols:
        if name_col in df_ship.columns:
            df_ship.rename(columns={name_col: '商品名'}, inplace=True)
            break
    if '商品名' not in df_ship.columns:
        df_ship['商品名'] = None  # 如果没有商品名列，也要占位，避免后续merge出错

    df_ship[date_col] = pd.to_datetime(df_ship[date_col], errors='coerce')
    df_ship = df_ship.dropna(subset=[date_col])
    df_ship['年月'] = df_ship[date_col].dt.month.astype(str) + '月'

    df_ship[code_col] = df_ship[code_col].astype(str).str.strip()
    df_ketsu["商品コード"] = df_ketsu["商品コード"].astype(str).str.strip()

    dummy_df = df_ship[[trade_col, code_col, '年月', '欠品理由']].dropna()
    dummy_df['欠品理由'] = dummy_df['欠品理由'].astype(str).str.split(',')
    dummy_df = dummy_df.explode('欠品理由')
    dummy_df['欠品理由'] = dummy_df['欠品理由'].str.strip()
    dummies = pd.get_dummies(dummy_df['欠品理由'], prefix='欠品理由', dtype=int)
    dummy_df = pd.concat([dummy_df[[trade_col, code_col, '年月']], dummies], axis=1)
    dummy_grouped = dummy_df.groupby([trade_col, code_col, '年月'], as_index=False).sum()
    dummy_grouped['欠品理由_合計'] = dummy_grouped.drop(columns=[trade_col, code_col, '年月']).sum(axis=1)
    dummy_cols = [col for col in dummy_grouped.columns if col.startswith('欠品理由_') and col != '欠品理由_合計']

    df_ship['出貨回数'] = 1
    ship_count_df = df_ship.groupby([code_col, '年月'], as_index=False)['出貨回数'].count()

    drop_cols = ['工場指示日', '確定日', 'Order No', '管理番号']
    df_ship = df_ship.drop(columns=[col for col in drop_cols if col in df_ship.columns])
    group_cols = [trade_col, code_col, '年月']
    numeric_cols = df_ship.select_dtypes(include=['number']).columns.tolist()
    non_numeric_cols = [col for col in df_ship.columns if col not in numeric_cols and col not in group_cols and col != date_col]
    agg_dict = {col: 'sum' for col in numeric_cols}
    for col in non_numeric_cols:
        agg_dict[col] = 'first'
    if '商品名' in df_ship.columns:
        agg_dict['商品名'] = 'first'

    df_ship_grouped = df_ship.groupby(group_cols, as_index=False).agg(agg_dict)

    df_merged = pd.merge(df_ship_grouped, df_ketsu, left_on=[code_col, '年月'], right_on=["商品コード", "年月"], how="left")
    
    # merge 后修正商品名字段（如果产生了 _x/_y）
    if '商品名_x' in df_merged.columns and '商品名_y' in df_merged.columns:
        df_merged['商品名'] = df_merged['商品名_x']
        df_merged.drop(columns=['商品名_x', '商品名_y'], inplace=True)
    elif '商品名_x' in df_merged.columns:
        df_merged.rename(columns={'商品名_x': '商品名'}, inplace=True)
    elif '商品名' not in df_merged.columns and '商品名_y' in df_merged.columns:
        df_merged.rename(columns={'商品名_y': '商品名'}, inplace=True)
    
    df_merged = pd.merge(df_merged, dummy_grouped, on=[trade_col, code_col, '年月'], how="left")
    df_merged = pd.merge(df_merged, ship_count_df, on=[code_col, '年月'], how="left")
    if not df_status.empty:
        df_status['商品コード'] = df_status['商品コード'].astype(str).str.strip()
        df_merged = pd.merge(df_merged, df_status[['商品コード', '商品状態']], on='商品コード', how='left')

    df_merged.fillna(0, inplace=True)
    df_merged['三倉庫欠品日数合計'] = (
        df_merged.get('0001倉庫欠品日数', 0) +
        df_merged.get('1001倉庫欠品日数', 0) +
        df_merged.get('2001倉庫欠品日数', 0)
    )

    # 出貨回数去重处理（保留 _x 版本）
    if '出貨回数_x' in df_merged.columns and '出貨回数_y' in df_merged.columns:
        df_merged['出貨回数'] = df_merged['出貨回数_x']
        df_merged.drop(columns=['出貨回数_x', '出貨回数_y'], inplace=True)

    # ✅ 指定最终输出列顺序（统一格式）
    dummy_cols = [col for col in df_merged.columns if col.startswith('欠品理由_') and col != '欠品理由_合計']

    ordered_cols = ['商品状態', '年月', '商品コード']
    for col in ['商品名', trade_col, '工場名']:
        if col in df_merged.columns:
            ordered_cols.append(col)


    # 出荷类数值列（非控制列）
    skip_cols = set(ordered_cols + dummy_cols + [
    '0001倉庫欠品日数', '1001倉庫欠品日数', '2001倉庫欠品日数',
    '三倉庫欠品日数合計', '出貨回数', '欠品理由_合計'])
    ship_numeric = [col for col in df_merged.columns if col not in skip_cols and df_merged[col].dtype != 'O']


    # 拼接最终顺序
    final_cols = (
    ordered_cols +
    ship_numeric +
    ['出貨回数', '欠品理由_合計', '0001倉庫欠品日数', '1001倉庫欠品日数', '2001倉庫欠品日数', '三倉庫欠品日数合計'] +
    dummy_cols)
    final_cols = [col for col in final_cols if col in df_merged.columns]
    df_final = df_merged[final_cols].drop_duplicates()

    st.subheader("統合結果データ")
    st.dataframe(df_final)

    df_sales_long = pd.DataFrame()
    if isinstance(df_sales_raw, pd.DataFrame) and not df_sales_raw.empty:
        try:
            id_vars = ['商品コード']
            value_vars = [col for col in df_sales_raw.columns if re.search(r'20\d{2}[/年\-]\d{1,2}', str(col))]
            df_sales_long = df_sales_raw.melt(id_vars=id_vars, value_vars=value_vars, var_name='年月', value_name='販売量')
            df_sales_long['商品コード'] = df_sales_long['商品コード'].astype(str).str.strip()
            df_sales_long['年月'] = df_sales_long['年月'].astype(str)
            df_sales_long['年月'] = pd.to_datetime(df_sales_long['年月'], errors='coerce')
            df_sales_long = df_sales_long.dropna(subset=['年月'])
            df_sales_long['年月'] = df_sales_long['年月'].dt.strftime('%Y/%m')

            補完列 = ['商品コード', '商品名', '貿易会社名', '工場名', '商品状態']
            df_sales_info = df_sales_raw[補完列].drop_duplicates(subset=['商品コード'], keep='first')
            df_sales_long = pd.merge(df_sales_long, df_sales_info, on='商品コード', how='left')

            desired_cols = ['商品状態', '年月', '商品コード', '商品名', '貿易会社名', '工場名', '販売量']
            df_sales_long = df_sales_long[[col for col in desired_cols if col in df_sales_long.columns]]

            st.subheader("📊 販売量データ（整形後）")
            st.dataframe(df_sales_long)

        except Exception as e:
            st.warning(f"販売量データの整形時にエラーが発生しました: {e}")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name="統合結果")
        if not df_sales_long.empty:
            df_sales_long.to_excel(writer, index=False, sheet_name="販売量データ")
    buf.seek(0)

    st.download_button(
        "Excelダウンロード（統合結果＋販売量）",
        data=buf,
        file_name="統合結果_状態付き_販売量付き.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.session_state.cleaned_df = df_final  # ✅ 传给分析页的数据（整合结果）
