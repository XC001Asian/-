import pandas as pd
import os
import re

# ==========================================
# 1. 設定區域
# ==========================================
INPUT_CSV = '1m.csv'       # 網銀下載的 csv 檔名
HISTORY_DB = 'history.xlsx'       # 歷史紀錄檔名
OUTPUT_FILE = '分類結果_待確認.xlsx' # 輸出的結果檔名

# --- 第一層邏輯：萬用代號規則 ---
UNIVERSAL_RULES = {
    r'55\d{2}60': '6C',
    r'551403': '14G',
    '89': 'B2-89',
    '90': 'B2-90',
    r'3\d{2}105': '10I',
    'B195': 'B1-95',
    r'53\d{2}82': '8A',
    r'55\d{2}92': '9F',
    r'55\d{2}91': '9E',
    r'55\d{2}30': '3C',
    r'5\d{2}103': '10G',
    r'55\d{2}41': '4D',
    r'5\d{2}121': '12E',
    '14B': '14B',
    '4E': '4E',
    '12J': '12J',
    '12B': '12B',
    r'5\d{2}141': '14E',
    r'5\d{2}142': '14F',
    r'3\d{2}112': '11A',
    '34': 'B6-34',
    r'55\d{2}43': '4F',
    '7BB627': '7BB627',
    r'3\d{2}125': '12I',
    r'3\d{2}140': '14C',
    '6D': '6D',
    r'55\d{2}70': '7D',
    '9J': '9J',
    '55號  11樓  D': '11D',
    'b6-36': 'B6-36',
    '953300047': '47',
    r'53\d{2}95': '9I',
    r'3\d{2}132': '13A',
    r'53\d{2}52': '5I',
    r'53\d{2}61': '6A',
    r'53\d{2}62': '6I',

}

# ==========================================
# 2. 讀取歷史資料庫
# ==========================================

def load_history(path):
    if not os.path.exists(path):
        print("尚未發現歷史資料庫，將略過第二層邏輯...")
        return {}
        
    print(f"正在讀取歷史資料庫: {path}")
    
    try:
        dfh = pd.read_excel(path, dtype=str, engine='openpyxl')
        
        if '存匯代號' not in dfh.columns or '歸屬人' not in dfh.columns:
            print("❌ 錯誤：歷史資料庫缺少 `存匯代號` 或 `歸屬人` 欄位！")
            return {}

        dfh['存匯代號'] = dfh['存匯代號'].astype(str).str.strip()
        dfh['歸屬人'] = dfh['歸屬人'].astype(str).str.strip()
        dfh['存匯代號_末六碼'] = dfh['存匯代號'].str[-6:]
        
        history_dict_local = dict(zip(dfh['存匯代號_末六碼'], dfh['歸屬人']))
        
        print(f"✅ 歷史資料庫讀取成功！共 {len(history_dict_local)} 筆記錄。")
        return history_dict_local
        
    except Exception as e:
        print(f"❌ 讀取歷史資料發生錯誤: {e}")
        return {}

history_dict = load_history(HISTORY_DB)

# ==========================================
# 3. 讀取網銀 CSV
# ==========================================
print(f"正在讀取網銀資料: {INPUT_CSV}")

try:
    SKIP_ROWS = 0
    df_new = pd.read_csv(INPUT_CSV, encoding='cp950', dtype=str, skiprows=SKIP_ROWS)
    df_new.columns = df_new.columns.str.strip()
    
    # 清理所有欄位 (移除 ="...")
    def _clean_str(x):
        if pd.isna(x): return x
        s = str(x).strip()
        if s.startswith('="') and s.endswith('"'): s = s[2:-1]
        if s.startswith('=') and not s.startswith('="'): s = s.lstrip('=')
        return s.strip(' "')

    df_new.columns = [ _clean_str(c) for c in df_new.columns ]
    for col in df_new.select_dtypes(include=['object']).columns:
        df_new[col] = df_new[col].apply(_clean_str)

    # 檢查必要欄位
    required_cols = ['存匯代號', '支出金額'] # 我們現在多需要一個欄位
    for col in required_cols:
        if col not in df_new.columns:
            print(f"❌ 錯誤：找不到 '{col}' 欄位，請檢查 CSV 內容。")
            print("系統讀到的欄位：", df_new.columns.tolist())
            exit()

except Exception as e:
    print(f"❌ 讀取 CSV 失敗: {e}")
    exit()

# ==========================================
# 4. 核心分類邏輯 (新增第0層邏輯)
# ==========================================

# 注意：這裡現在接收的是整行資料 (row)，而不只是代號 (code)
def classify_logic(row):
    
    # --- 第 0 層：檢查是不是支出 ---
    # 取得支出金額，並嘗試轉成數字判斷
    expense_str = str(row.get('支出金額', '0')).replace(',', '').strip()
    try:
        # 如果是空的或 '-'，當作 0
        if not expense_str or expense_str == '-':
            expense_val = 0.0
        else:
            expense_val = float(expense_str)
            
        # 如果支出大於 0，直接回傳不用分類
        if expense_val > 0:
            return '支出(不需分類)', '系統(支出忽略)'
    except:
        pass # 如果轉換數字失敗，就繼續往下跑原本的邏輯
    
    # ======================================
    # 以下為原本的邏輯 (針對存入款項)
    # ======================================
    code = str(row.get('存匯代號', '')).strip()
    
    if not code or code.lower() == 'nan':
        return '未知(無代號)', '需人工'

    # --- 第一順位：查詢歷史資料庫 (比對末六碼) ---
    short_code = code[-6:] 
    if short_code in history_dict:
        return history_dict[short_code], '系統(歷史紀錄-末六碼)'

    # --- 第二順位：萬用代號規則 (比對 Regex) ---
    for pattern, owner in UNIVERSAL_RULES.items():
        if re.search(pattern, code):
            return owner, '系統(萬用代號)'

    # --- 第三順位：人工 ---
    return '待人工確認', '需人工'


# ==========================================
# 5. 執行與輸出
# ==========================================

print("正在進行分類...")

# 【關鍵修改】 axis=1 代表我們要「一行一行」送進去函式，而不是「一格一格」
df_new[['歸屬人', '判斷來源']] = df_new.apply(
    lambda row: pd.Series(classify_logic(row)), axis=1
)

df_manual = df_new[df_new['判斷來源'] == '需人工']
df_auto = df_new[df_new['判斷來源'] != '需人工']

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df_new.to_excel(writer, sheet_name='總表', index=False)
        if not df_manual.empty:
            df_manual.to_excel(writer, sheet_name='需人工確認清單', index=False)
            
    print("="*30)
    print(f"處理完成！檔案已儲存為：{OUTPUT_FILE}")
    print(f"總筆數：{len(df_new)}")
    print(f"系統自動/忽略支出：{len(df_auto)} 筆")
    print(f"需要您手工確認：{len(df_manual)} 筆")
    print("="*30)
except PermissionError:
    print(f"❌ 無法存檔！請先關閉 '{OUTPUT_FILE}' 檔案後再試一次。")
    
# 首次建立範本
if not os.path.exists(HISTORY_DB):
    sample_db = pd.DataFrame({'存匯代號': ['範例123456', '範例889900'], '歸屬人': ['客戶A', '客戶B']})
    sample_db.to_excel(HISTORY_DB, index=False)
    print(f"已建立歷史資料庫範本：{HISTORY_DB}")