import pandas as pd
import glob
import os
import sys

# ================= 設定區 =================

# 1. 取得腳本所在的基本路徑 (Base Path)
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
elif '__file__' in locals():
    base_path = os.path.dirname(os.path.abspath(__file__))
else:
    base_path = os.getcwd()

# 2. 指定資料來源資料夾名稱 (根據你的描述是這個)
target_folder_name = "data-2015-2024"
input_path = os.path.join(base_path, target_folder_name)

# 3. 設定輸出檔名 (存放在腳本旁邊，也就是 datastream 資料夾)
final_excel_name = "all-countries.xlsx"
output_excel_path = os.path.join(base_path, final_excel_name)

# 4. 暫存檔與紀錄檔
final_csv_name = "all-countries.csv"
output_csv_path = os.path.join(base_path, final_csv_name)

log_file_name = "processed_log.txt"
log_path = os.path.join(base_path, log_file_name)

# ================= 檢查舊檔案 =================
existing_files = []
for f in [output_excel_path, log_path, output_csv_path]:
    if os.path.exists(f):
        existing_files.append(f)

if existing_files:
    print("偵測到以下舊檔案可能是上次執行時產生的：")
    for f in existing_files:
        print(f" - {f}")
    
    choice = input("你想刪除這些檔案並重新執行嗎？(y/n): ").strip().lower()
    if choice == 'y':
        for f in existing_files:
            try:
                os.remove(f)
                print(f"已刪除: {f}")
            except Exception as e:
                print(f"[錯誤] 無法刪除 {f}: {e}")
    else:
        print("程式已停止，保留舊檔案以避免覆寫。")
        sys.exit()


# ==========================================

print(f"程式位置: {base_path}")
print(f"正在搜尋資料夾: {input_path}")
print("-" * 30)

# 檢查資料夾是否存在
if not os.path.exists(input_path):
    print(f"[錯誤] 找不到資料夾：{target_folder_name}")
    print(f"請確認你的目錄結構如下：")
    print(f"{base_path}\\")
    print(f"  └── {target_folder_name}\\ (Excel要放在這裡)")
    input("按 Enter 離開...")
    sys.exit()

# 讀取「已完成清單」
processed_files = set()
if os.path.exists(log_path):
    with open(log_path, "r", encoding="utf-8") as f:
        processed_files = set(line.strip() for line in f)

# 抓取 target_folder_name 內所有的 .xlsx
all_files = glob.glob(os.path.join(input_path, "*.xlsx"))
print(f"發現 {len(all_files)} 個 Excel 檔案。")

count = 0
for filename in all_files:
    file_basename = os.path.basename(filename)
    
    # === 過濾區 ===
    if file_basename.startswith("~$"): continue
    if file_basename in processed_files: continue
    # =============

    print(f"正在合併: {file_basename}: {len(df.columns)} 欄")
    
    try:
        df = pd.read_excel(filename, dtype=str)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # 去掉多餘的空白欄
        
        # 寫入 CSV (存放在外面那一層，避免汙染資料夾)
        file_exists = os.path.isfile(output_csv_path)
        df.to_csv(output_csv_path, mode='a', index=False, header=not file_exists, encoding='utf-8-sig')
        
        # 寫入 Log
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(file_basename + "\n")
            
        count += 1
        
    except Exception as e:
        print(f"[錯誤] 讀取 {file_basename} 失敗: {e}")

print("-" * 30)
print(f"本次新增合併 {count} 個檔案。")

# ================= 轉存 Excel =================
if os.path.exists(output_csv_path):
    print(f"即將建立最終檔案: {final_excel_name}，可能會花幾分鐘...")
    
    user_input = input("是否要轉存為 Excel？(y/n): ").strip().lower()
    if user_input == "y":
        try:
            df_final = pd.read_csv(output_csv_path, dtype=str)  # 指定型態為字串，避免 DtypeWarning
            df_final.to_excel(output_excel_path, index=False)
            print(f"\n★ 成功！檔案位置: {output_excel_path}")
        except Exception as e:
            print(f"轉存 Excel 失敗: {e}")
    else:
        print("跳過轉存 Excel，只輸出 CSV 檔案。")
else:
    if count == 0:
        print("沒有新檔案需要合併。")
