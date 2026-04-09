import streamlit as st
import pandas as pd
import numpy as np
import io
import os

st.set_page_config(page_title="継続MAS予算登録用アプリ", layout="wide")

def load_granular_data(csv_file):
    base_dir = os.path.dirname(__file__)
    excel_file = os.path.join(base_dir, "mas_template.xlsx")
    
    df_excel = pd.read_excel(excel_file)
    df_excel['Cat'] = df_excel['利益管理表科目'].bfill()
    df_excel['Cat'] = df_excel['Cat'].fillna('その他')
    
    df_acc = df_excel[df_excel['勘定科目コード'].notna() & df_excel['勘定科目名'].notna()].copy()
    df_acc['勘定科目コード'] = df_acc['勘定科目コード'].astype(int)
    
    def get_group(code):
        if 4000 <= code < 5000: return '売上'
        elif 5400 <= code < 5500: return '製造原価'
        elif 5000 <= code < 6000: return '売上原価・変動費'
        elif 6000 <= code < 7000: return '固定費'
        elif 7000 <= code < 8000: return '営業外'
        return 'その他'
        
    df_acc['Group'] = df_acc['勘定科目コード'].apply(get_group)
    
    try:
        df_csv = pd.read_csv(csv_file, encoding='utf-8-sig')
    except:
        df_csv = pd.read_csv(csv_file, encoding='shift_jis')
        
    df_csv['Account Name_CSV'] = df_csv['勘定科目名'].astype(str).str.strip()
    df_csv['CSV Code'] = pd.to_numeric(df_csv['勘定科目コード'], errors='coerce')
    df_csv = df_csv[df_csv['CSV Code'].notna()]
    df_csv['CSV Code'] = df_csv['CSV Code'].astype(int)
    df_csv['Actual'] = pd.to_numeric(df_csv['当月'], errors='coerce').fillna(0).astype(int)
    
    merged = pd.merge(df_acc, df_csv[['CSV Code', 'Account Name_CSV', 'Actual']], 
                      left_on='勘定科目コード', right_on='CSV Code', how='left')
                      
    df_acc_filtered = merged[~((merged['勘定科目コード'] % 100 == 0) | merged['勘定科目コード'].isin([5410, 5420, 5430]))].copy()
    
    df_acc_filtered['Final Name'] = df_acc_filtered['Account Name_CSV'].fillna(df_acc_filtered['勘定科目名'])
    df_acc_filtered.loc[df_acc_filtered['勘定科目コード'].isin([4111, 4112]), 'Final Name'] = '売上高'
    df_acc_filtered.loc[df_acc_filtered['勘定科目コード'] == 6232, 'Final Name'] = ''
    
    df_acc_filtered['勘定科目名'] = df_acc_filtered['Final Name']
    df_acc_filtered['Actual'] = df_acc_filtered['Actual'].fillna(0).astype(int)
    
    return df_acc_filtered, df_acc['Cat'].unique()

def calculate_budget1(df, target_profit=None):
    res = df.copy()
    
    totals = res.groupby('Group')['Actual'].sum()
    sales = totals.get('売上', 0)
    vc = totals.get('売上原価・変動費', 0) + totals.get('製造原価', 0)
    fc = totals.get('固定費', 0)
    non_op = totals.get('営業外', 0)
    
    actual_profit = sales - vc - fc + non_op
    vcr = vc / sales if sales != 0 else 0
    
    scenario_desc = ""
    if target_profit is None:
        if actual_profit < 0:
            target_profit = 100000
            scenario_desc = "赤字回復逆算 (利益10万目標)"
        else:
            required_sales = sales * 1.05
            target_profit = required_sales - (required_sales * vcr) - fc + non_op
            scenario_desc = "黒字成長 (売上5%UP)"
            
    required_sales = (target_profit - non_op + fc) / (1 - vcr) if vcr < 1 else sales
    multiplier = required_sales / sales if sales != 0 else 1
    
    budget1_vals = []
    for idx, row in res.iterrows():
        g = row['Group']
        val = row['Actual']
        if g in ['売上', '売上原価・変動費', '製造原価']:
            budget1_vals.append(int(round(val * multiplier)))
        else:
            budget1_vals.append(int(val))
            
    res['Budget1'] = budget1_vals
    res['Budget2'] = res['Budget1']
    
    return res, target_profit, scenario_desc

def render_group_editor(df, cat_name):
    subset = df[df['Cat'] == cat_name].copy()
    if subset.empty:
        return pd.DataFrame(), pd.DataFrame()
        
    disp = subset[['勘定科目名', 'Actual', 'Budget1', 'Budget2']].copy()
    disp.columns = ['科目名', '当期実績', '予算１(自動)', '予算２(手入力)']
    
    config = {
        '科目名': st.column_config.TextColumn("科目名"),
        '当期実績': st.column_config.NumberColumn(format="%,.0f"),
        '予算１(自動)': st.column_config.NumberColumn(format="%,.0f"),
        '予算２(手入力)': st.column_config.NumberColumn(format="%,.0f")
    }
    
    header_col = st.empty()
    edited = st.data_editor(
        disp,
        disabled=['当期実績', '予算１(自動)'],
        num_rows="dynamic",
        hide_index=True,
        column_config=config,
        use_container_width=True,
        key=f"editor_b1_{cat_name}"
    )
    edited['当期実績'] = edited['当期実績'].fillna(0).astype(int)
    edited['予算１(自動)'] = edited['予算１(自動)'].fillna(0).astype(int)
    edited['予算２(手入力)'] = edited['予算２(手入力)'].fillna(0).astype(int)
    edited['科目名'] = edited['科目名'].fillna("新規科目")
    
    # Store category and group mapping for reconstruction
    group_map = subset.iloc[0]['Group'] if not subset.empty else "その他"
    
    df_out = edited.copy()
    df_out['Cat'] = cat_name
    df_out['Group'] = group_map
    df_out.rename(columns={'科目名': '勘定科目名'}, inplace=True)
    
    subtotal_act = edited['当期実績'].sum()
    subtotal_b2 = edited['予算２(手入力)'].sum()
    header_col.markdown(f"**▪ {cat_name}** 　*(小計: 実績 ¥{subtotal_act:,.0f} / 予算2 ¥{subtotal_b2:,.0f})*")
    
    return edited, df_out

def render_5y_group_editor(yr1_df, cat_name, growth_rate):
    subset = yr1_df[yr1_df['Cat'] == cat_name].copy()
    len_subset = len(subset)
    if len_subset == 0:
        return pd.DataFrame(), pd.DataFrame()
        
    m = np.zeros((len_subset, 5), dtype=int)
    m[:, 0] = subset['予算２(手入力)'].values
    
    group_name = subset.iloc[0]['Group'] if not subset.empty else "その他"
    is_growing_group = group_name in ['売上', '売上原価・変動費', '製造原価']
    mult = (1 + growth_rate) if is_growing_group else 1.0
    
    for i in range(1, 5):
        m[:, i] = np.round(m[:, i-1] * mult)
        
    disp = pd.DataFrame(m, columns=['第1期', '第2期', '第3期', '第4期', '第5期'])
    disp.insert(0, '科目名', subset['勘定科目名'].values)
    
    config = {col: st.column_config.NumberColumn(format="%,.0f") for col in ['第1期', '第2期', '第3期', '第4期', '第5期']}
    config['科目名'] = st.column_config.TextColumn("科目名")
    
    header_col = st.empty()
    edited = st.data_editor(
        disp,
        num_rows="dynamic",
        hide_index=True,
        disabled=[],
        column_config=config,
        use_container_width=True,
        key=f"editor_5y_{cat_name}"
    )
    for col in ['第1期', '第2期', '第3期', '第4期', '第5期']:
        edited[col] = edited[col].fillna(0).astype(int)
    edited['科目名'] = edited['科目名'].fillna("新規科目")
    
    df_out = edited.copy()
    df_out['Cat'] = cat_name
    df_out['Group'] = group_name
    df_out.rename(columns={'科目名': '勘定科目名'}, inplace=True)
    
    s_y1 = edited['第1期'].sum()
    s_y2 = edited['第2期'].sum()
    s_y3 = edited['第3期'].sum()
    header_col.markdown(f"**▪ {cat_name}** 　*(小計: 第1期 ¥{s_y1:,.0f} / 第2期 ¥{s_y2:,.0f} / 第3期 ¥{s_y3:,.0f} ...)*")
    
    return edited, df_out

def main():
    st.title("継続MAS予算登録用アプリ")
    
    st.sidebar.header("Data Upload")
    csv_file = st.sidebar.file_uploader("勘定科目残高.csv をアップロード", type=['csv'])
    
    if csv_file:
        raw_df, cats_order = load_granular_data(csv_file)
        
        if 'target_profit_override' not in st.session_state:
            st.session_state['target_profit_override'] = None
            
        b1_df, curr_target, scenario_desc = calculate_budget1(raw_df, st.session_state['target_profit_override'])
        
        tab1, tab2, tab4 = st.tabs(["実績 ＆ 予算（1・2）", "5カ年計画", "借入金・C/F予測"])
        
        with tab1:
            st.header("1. 前期実績 と 予算シミュレーション")
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.info("💡 **予算1逆算の整合性**: ここで「目標経常利益」を変更すると、以下の『売上』『変動費』の全科目が現在の比率のまま逆算配分されます。")
            with col2:
                new_target = st.number_input("予算１ 目標経常利益", value=int(curr_target), step=10000)
                if new_target != int(curr_target) and st.session_state['target_profit_override'] != new_target:
                    st.session_state['target_profit_override'] = new_target
                    st.rerun()
            
            st.subheader("【明細入力】予算２ シミュレータ")
            
            b1_sum_sales = 0; b1_sum_vc = 0; b1_sum_fc = 0; b1_sum_nonop = 0
            b2_sum_sales = 0; b2_sum_vc = 0; b2_sum_fc = 0; b2_sum_nonop = 0
            
            recon_b1 = []
            
            for cat in cats_order:
                ed_df, out_df = render_group_editor(b1_df, cat)
                if not out_df.empty:
                    recon_b1.append(out_df)
                    g = out_df.iloc[0]['Group']
                    v1 = out_df['予算１(自動)'].sum()
                    v2 = out_df['予算２(手入力)'].sum()
                    
                    if g == '売上':
                        b1_sum_sales += v1; b2_sum_sales += v2
                    elif g in ['売上原価・変動費', '製造原価']:
                        b1_sum_vc += v1; b2_sum_vc += v2
                    elif g == '固定費':
                        b1_sum_fc += v1; b2_sum_fc += v2
                    elif g == '営業外':
                        b1_sum_nonop += v1; b2_sum_nonop += v2
                
                if cat == "他の変動費":
                    b2_vcr = b2_sum_vc / b2_sum_sales if b2_sum_sales else 0
                    b2_mp_val = b2_sum_sales - b2_sum_vc
                    st.info(f"💡 **限界利益**: ¥{b2_mp_val:,.0f} （限界利益率: {1-b2_vcr:.1%} / 変動費率: {b2_vcr:.1%}）")
                    
                if cat == "保険料・修繕費":
                    b2_prof_val = (b2_sum_sales - b2_sum_vc) - b2_sum_fc + b2_sum_nonop
                    st.info(f"💡 **経常利益**: ¥{b2_prof_val:,.0f}")
            
            if recon_b1:
                dynamic_b1_df = pd.concat(recon_b1, ignore_index=True)
            else:
                dynamic_b1_df = pd.DataFrame()
            
            b1_prof = b1_sum_sales - b1_sum_vc - b1_sum_fc + b1_sum_nonop
            b1_vcr = b1_sum_vc / b1_sum_sales if b1_sum_sales else 0
            b1_mp = b1_sum_sales - b1_sum_vc
            
            b2_prof = b2_sum_sales - b2_sum_vc - b2_sum_fc + b2_sum_nonop
            b2_vcr = b2_sum_vc / b2_sum_sales if b2_sum_sales else 0
            b2_mp = b2_sum_sales - b2_sum_vc
            
            st.success(f"**予算1 逆算結果**: 【売上高】 ¥{b1_sum_sales:,.0f} － 【変動費】 ¥{b1_sum_vc:,.0f} （変動費率: {b1_vcr:.1%}） ＝ **【限界利益】 ¥{b1_mp:,.0f}**<br>→ 限界利益 － 【固定費】 ¥{b1_sum_fc:,.0f} ＋ 【営業外損益】 ¥{b1_sum_nonop:,.0f} ＝ **【経常利益】 ¥{b1_prof:,.0f}**", icon="🎯")
            st.info(f"**現在の予測2 シミュレーション集計**: 【売上高】 ¥{b2_sum_sales:,.0f} － 【変動費】 ¥{b2_sum_vc:,.0f} （変動費率: {b2_vcr:.1%}） ＝ **【限界利益】 ¥{b2_mp:,.0f}** <br>→ 限界利益 － 【固定費】 ¥{b2_sum_fc:,.0f} ＋ 【営業外損益】 ¥{b2_sum_nonop:,.0f} ＝ **【経常利益】 ¥{b2_prof:,.0f}**", icon="📊")

        with tab2:
            st.header("2. 5カ年計画 (明細シミュレータ)")
            st.info("💡 1年目(第1期)は「予算2」を継承しています。ここで各期の数値を個別に書き換えることも可能です。")
            growth_rate = st.number_input("自動配分用の売上・変動費 年間成長率 (%)", value=2.0, step=0.1) / 100
            
            years = []
            recon_5y = []
            
            # Use dynamic chunks for accurate summation
            for cat in cats_order:
                ed_df, out_df = render_5y_group_editor(dynamic_b1_df, cat, growth_rate)
                if not out_df.empty:
                    recon_5y.append(out_df)
            
            if recon_5y:
                full_5y_df = pd.concat(recon_5y, ignore_index=True)
            else:
                full_5y_df = pd.DataFrame()
            
            for i in range(1, 6):
                col_name = f'第{i}期'
                if not full_5y_df.empty:
                    sales = full_5y_df[full_5y_df['Group'] == '売上'][col_name].sum()
                    vc = full_5y_df[full_5y_df['Group'].isin(['売上原価・変動費', '製造原価'])][col_name].sum()
                    fc = full_5y_df[full_5y_df['Group'] == '固定費'][col_name].sum()
                    nonop = full_5y_df[full_5y_df['Group'] == '営業外'][col_name].sum()
                    
                    deprec_rows = full_5y_df[full_5y_df['勘定科目名'].str.contains('減価償却', na=False)]
                    deprec = int(deprec_rows[col_name].sum()) if not deprec_rows.empty else 0
                else:
                    sales = vc = fc = nonop = deprec = 0
                    
                profit = (sales - vc) - fc + nonop
                years.append({
                    '売上高': sales, '変動費計': vc, '限界利益': sales - vc, 
                    '変動費率': vc/sales if sales else 0,
                    '固定費計': fc, '営業外損益': nonop, '経常利益': profit,
                    '減価償却費': deprec
                })
                
            df_5y_summary = pd.DataFrame(years, index=[f"第{i}期" for i in range(1, 6)])
            
        with tab4:
            st.header("3. 借入金返済 と キャッシュ・フロー予測")
            
            colA, colB, colC = st.columns(3)
            ar_days = colA.number_input("売上債権 回転日数 (日)", value=30.0)
            inv_days = colB.number_input("棚卸資産 回転日数 (日)", value=15.0)
            ap_days = colC.number_input("買入債務 回転日数 (日)", value=30.0)
            
            bs_list = []
            for i, row in df_5y_summary.iterrows():
                sales = row['売上高']
                vc = row['変動費計']
                
                ar = int(round((sales * ar_days) / 365))
                inv = int(round((vc * inv_days) / 365))
                ap = int(round((vc * ap_days) / 365))
                
                bs_list.append({
                    "売上債権残高": ar,
                    "棚卸資産残高": inv,
                    "買入債務残高": ap,
                })
                
            bs_df = pd.DataFrame(bs_list, index=[f"第{i}期" for i in range(1, 6)])
            st.write("▼ 予測貸借対照表 (B/S) ピックアップ")
            st.dataframe(bs_df.T.style.format("{:,.0f}"), use_container_width=True)
            
            st.divider()
            
            loan_input = pd.DataFrame([
                {"借入先": "銀行A", "現在残高": 10000000, "年間返済額": 2000000, "残り期間(年)": 5, "第1期・新規借入": 0},
                {"借入先": "銀行B", "現在残高": 5000000, "年間返済額": 1000000, "残り期間(年)": 5, "第1期・新規借入": 0}
            ])
            
            edit_conf = {
                "現在残高": st.column_config.NumberColumn(format="%,.0f"),
                "年間返済額": st.column_config.NumberColumn(format="%,.0f"),
                "残り期間(年)": st.column_config.NumberColumn(format="%,.0f"),
                "第1期・新規借入": st.column_config.NumberColumn(format="%,.0f"),
            }
            edited_loans = st.data_editor(loan_input, num_rows="dynamic", column_config=edit_conf, use_container_width=True)
            
            repayments_per_year = []
            for i in range(5):
                yr_rep = 0
                for _, loan in edited_loans.iterrows():
                    if loan["残り期間(年)"] > i:
                        yr_rep += loan["年間返済額"]
                repayments_per_year.append(int(yr_rep))
                
            act_ar = raw_df[raw_df['勘定科目名'].str.contains('売掛金|受取手形', na=False)]['Actual'].sum()
            act_inv = raw_df[raw_df['勘定科目名'].str.contains('商品|製品|仕掛品|材料', na=False)]['Actual'].sum()
            act_ap = raw_df[raw_df['勘定科目名'].str.contains('買掛金|支払手形', na=False)]['Actual'].sum()
            
            prev_ar = int(act_ar)
            prev_inv = int(act_inv)
            prev_ap = int(act_ap)
            
            cf_list = []
            for i in range(5):
                curr = df_5y_summary.iloc[i]
                
                net_profit = int(curr['経常利益'])
                depreciation = int(curr['減価償却費'])
                
                operating_cf = net_profit + depreciation
                
                new_loans = int(edited_loans["第1期・新規借入"].sum()) if i == 0 else 0
                financial_cf = new_loans - repayments_per_year[i]
                fcf = operating_cf + financial_cf
                
                cf_list.append({
                    "経常利益": net_profit,
                    "減価償却費 (+加算)": depreciation,
                    "営業C/F": operating_cf,
                    "財務C/F (新規借入-返済)": financial_cf,
                    "合計 フリーC/F": fcf
                })

            cf_df = pd.DataFrame(cf_list, index=[f"第{i}期" for i in range(1, 6)])
            st.divider()
            st.write("▼ キャッシュ・フロー予測")
            st.dataframe(cf_df.T.style.format("{:,.0f}"), use_container_width=True)
            st.bar_chart(cf_df["合計 フリーC/F"])
            
        # PDF EXPORT HOOK (to be integrated)
        st.divider()
        st.subheader("PDF ダウンロード")
        import mas_pdf_generator
        try:
            pdf_bytes = mas_pdf_generator.generate_pdf(full_5y_df, df_5y_summary, bs_df, cf_df)
            st.download_button(label="📄 全明細・5カ年計画をPDFでダウンロード", data=pdf_bytes, file_name="MAS_5Year_Plan.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"PDF Output Error: {e}")

    else:
        st.info("左側のサイドバーから「勘定科目残高.csv」をアップロードしてください。")

if __name__ == '__main__':
    main()
