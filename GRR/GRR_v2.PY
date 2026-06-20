import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MaxNLocator
import os
import sys
import matplotlib
import glob


matplotlib.use('Agg')

def find_coords(df, keyword):
    """關鍵詞定位邏輯"""
    search_rows, search_cols = min(df.shape[0], 150), min(df.shape[1], 50)
    for r in range(search_rows):
        for c in range(search_cols):
            cell_val = str(df.iloc[r, c]).strip()
            if keyword.lower() in cell_val.lower():
                return r, c
    return None, None

def process_single_csv(file_path):
    """處理單個 CSV，返回結果列表並生成 PDF/PNG"""
    file_path = file_path.strip().strip('"').strip("'")
    if not file_path.lower().endswith('.csv'):
        return []
        
    output_dir = file_path.replace(".csv", "_MSA_Plots")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    results_list = []

    try:
        raw_df = pd.read_csv(file_path, header=None, dtype=str, low_memory=False)
        
        sn_r, sn_c = find_coords(raw_df, "SerialNumber")
        para_r, para_c = find_coords(raw_df, "Parametric")
        usl_r, _ = find_coords(raw_df, "Upper Limit")
        lsl_r, _ = find_coords(raw_df, "Lower Limit")

        if sn_r is None or para_r is None:
            print(f"      [Skip] {os.path.basename(file_path)}: 標頭定位失敗。")
            return []

        name_row = para_r + 1
        test_indices = [c for c in range(para_c, raw_df.shape[1]) if pd.notna(raw_df.iloc[name_row, c])]
        
        data_start_row = 7
        for r in range(7, len(raw_df)):
            try:
                first_val = str(raw_df.iloc[r, test_indices[0]]).strip()
                if first_val and not pd.isna(first_val):
                    float(first_val)
                    data_start_row = r
                    break
            except (ValueError, IndexError): continue

        output_pdf = file_path.replace(".csv", "_MSA_Report.pdf")
        
        with PdfPages(output_pdf) as pdf:
            for col in test_indices:
                test_name = str(raw_df.iloc[name_row, col]).strip().replace("/", "_")
                df_sub = raw_df.iloc[data_start_row:, [sn_c, col]].copy()
                df_sub.columns = ['SN', 'Value']
                df_sub['Value'] = pd.to_numeric(df_sub['Value'], errors='coerce')
                clean_data = df_sub.dropna(subset=['Value']).copy()
                
                if clean_data.empty: continue

                clean_data['Trial'] = clean_data.groupby('SN').cumcount() + 1
                u_limit = pd.to_numeric(raw_df.iloc[usl_r, col], errors='coerce') if usl_r is not None else np.nan
                l_limit = pd.to_numeric(raw_df.iloc[lsl_r, col], errors='coerce') if lsl_r is not None else np.nan
                stds = clean_data.groupby('SN')['Value'].std()
                sigma_error = np.sqrt((stds**2).mean())
                sigma_part = clean_data.groupby('SN')['Value'].mean().std()
                tolerance = u_limit - l_limit
                ptr = (6 * sigma_error / tolerance * 100) if (tolerance > 0) else 0
                result = "PASS" if ptr < 30 else "FAIL"

                results_list.append({
                    "Source_File": os.path.basename(file_path),
                    "Test_Item": test_name,
                    "LSL": l_limit,
                    "USL": u_limit,
                    "PTR(%)": round(ptr, 2),
                    "Result": result,
                    "Sigma_Part": round(sigma_part, 6),
                    "Sigma_Error": round(sigma_error, 6)
                })

                fig = plt.figure(figsize=(11, 13))
                gs = fig.add_gridspec(3, 1, height_ratios=[1, 1, 0.3])
                ax1 = fig.add_subplot(gs[0])
                ax2 = fig.add_subplot(gs[1])
                ax_tbl = fig.add_subplot(gs[2])
                ax_tbl.axis('off')

                sns.lineplot(data=clean_data, x='Trial', y='Value', hue='SN', marker='o', ax=ax1, palette='tab20', legend=False)
                ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
                sns.boxplot(data=clean_data, x='SN', y='Value', ax=ax2, color='#f7f7f7', showfliers=False)
                sns.stripplot(data=clean_data, x='SN', y='Value', ax=ax2, color='blue', alpha=0.4, size=4)
                plt.setp(ax2.get_xticklabels(), rotation=45, ha='right', fontsize=8)

                for ax in [ax1, ax2]:
                    if not np.isnan(u_limit): ax.axhline(u_limit, color='red', linestyle='--', linewidth=1.5)
                    if not np.isnan(l_limit): ax.axhline(l_limit, color='red', linestyle='--', linewidth=1.5)

                table_data = [
                    ["Lower Limit", "Upper Limit", "PTR (%)", "Result", "σ_part", "σ_error"],
                    [f"{l_limit:.4f}", f"{u_limit:.4f}", f"{ptr:.2f}%", result, f"{sigma_part:.6f}", f"{sigma_error:.6f}"]
                ]
                tbl = ax_tbl.table(cellText=table_data, loc='center', cellLoc='center', colWidths=[0.15]*6)
                tbl.auto_set_font_size(False)
                tbl.set_fontsize(10)
                tbl.scale(1.2, 2.5)

                res_color = '#90EE90' if result == "PASS" else '#FFB6C1'
                for (row, col_idx), cell in tbl.get_celld().items():
                    if row == 0:
                        cell.set_facecolor('#E6E6E6') # 表頭灰色
                    else:
                        cell.set_facecolor(res_color) # 數據列根據結果變色

                plt.suptitle(f"MSA ANALYSIS REPORT: {test_name}", fontsize=16, fontweight='bold', y=0.98)
                plt.tight_layout(rect=[0, 0.05, 1, 0.95])
                pdf.savefig(fig)
                fig.savefig(os.path.join(output_dir, f"{test_name}_MSA.png"), dpi=150)
                plt.close(fig)
                
        return results_list
    except Exception as e:
        print(f"      [Error] {os.path.basename(file_path)}: {e}")
        return []

def generate_excel_summary(data_list, output_folder):
    """生成帶格式的 Excel 總表"""
    if not data_list: return
    excel_file = os.path.join(output_folder, "MSA_Summary_Report.xlsx")
    df = pd.DataFrame(data_list)
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Summary')
    workbook, worksheet = writer.book, writer.sheets['Summary']
    
    green_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    red_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    row_count = len(df)
    worksheet.conditional_format(1, 5, row_count, 5, {'type': 'cell', 'criteria': 'equal to', 'value': '"PASS"', 'format': green_fmt})
    worksheet.conditional_format(1, 5, row_count, 5, {'type': 'cell', 'criteria': 'equal to', 'value': '"FAIL"', 'format': red_fmt})
    worksheet.conditional_format(1, 4, row_count, 4, {'type': 'cell', 'criteria': '>', 'value': 30, 'format': red_fmt})

    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, column_len)
    writer.close()
    print(f"\n>>> [Success] 彩色匯總表已生成: {excel_file}")

if __name__ == "__main__":
    print("========================================")
    print("      GRR 批量生成    ")
    print("========================================")
    
    input_path = input("\n>>> 請將【資料夾】或【單個 CSV】拖入，按回車：").strip().strip('"').strip("'")
    all_summary_data = []
    
    if os.path.isdir(input_path):
        csv_files = glob.glob(os.path.join(input_path, "*.csv"))
        for i, f in enumerate(csv_files, 1):
            print(f"[{i}/{len(csv_files)}] 處理中: {os.path.basename(f)}")
            all_summary_data.extend(process_single_csv(f))
        generate_excel_summary(all_summary_data, input_path)
    elif os.path.isfile(input_path) and input_path.lower().endswith('.csv'):
        print(f"正在處理單一文件: {os.path.basename(input_path)}")
        all_summary_data = process_single_csv(input_path)
        generate_excel_summary(all_summary_data, os.path.dirname(input_path))
    else:
        print("\n>>> [Error] 無效路徑！")

    input("\n所有任務已完成，按回車退出...")
