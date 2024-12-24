import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from datetime import datetime, timedelta
import urllib.parse

# グローバル変数として設定
base_url = "https://www.slorepo.com/prefecture/?prefecture="

# WebDriverを初期化する関数
def init_driver():
    chrome_driver_path = r"D:\Users\Documents\python\chromedriver.exe"  # ここに正しいパスを指定
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # ヘッドレスモードで実行
    driver = webdriver.Chrome(service=Service(chrome_driver_path), options=options)
    return driver

# 選択された都道府県の店名リストを取得する関数
def fetch_store_names(prefecture):
    driver = init_driver()
    
    # URLエンコードを使って都道府県名をURLに組み込む
    prefecture_encoded = urllib.parse.quote(prefecture)
    prefecture_url = f"{base_url}{prefecture_encoded}"
    
    try:
        driver.get(prefecture_url)
        # ページソースを取得して解析
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # '更新中のホール' テキストを含む <strong> タグを探す
        update_halls_tag = soup.find('strong', string=re.compile(r'更新中のホール'))
        store_names = []

        if update_halls_tag:
            table = update_halls_tag.find_next('table')
            if table:
                rows = table.find_all('tr')
                for row in rows[1:]:  # ヘッダーをスキップ
                    cols = row.find_all('td')
                    if cols:
                        store_name = cols[0].text.strip()
                        store_url = cols[0].find('a')['href']
                        store_names.append((store_name, store_url))
            else:
                print("テーブルが見つかりませんでした。")
        else:
            print("更新中のホールタグが見つかりませんでした。")
    
    finally:
        driver.quit()
    
    return store_names

# 店名リストの更新
def update_store_dropdown():
    selected_prefecture = prefecture_var.get()
    if selected_prefecture:
        store_names_urls = fetch_store_names(selected_prefecture)
        store_names = [name for name, url in store_names_urls]
        if store_names:
            store_dropdown['values'] = store_names
            store_dropdown.set('')  # プルダウンのリセット
        else:
            messagebox.showinfo("情報", "該当する店名がありませんでした。")
    else:
        messagebox.showerror("エラー", "都道府県を選択してください。")

# 開始日を指定する関数
def get_valid_date():
    date_str = date_entry.get()
    try:
        date = datetime.strptime(date_str, "%Y%m%d")
        # 未来の日付はエラー
        if date > datetime.now():
            messagebox.showerror("エラー", "未来の日付は指定できません。")
            return None
        return date
    except ValueError:
        messagebox.showerror("エラー", "正しい日付形式で入力してください (YYYYMMDD)")
        return None

# URL生成
def generate_url(base_url, store_url, date):
    formatted_date = date.strftime("%Y%m%d")
    return f"{store_url}{formatted_date}/"

# ページスクレイピング処理
def scrape_page(driver, url, date_formatted):
    driver.get(url)
    all_data = []
    
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "strong")))
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        
        if "お探しのページは見つかりませんでした。" in html:
            print(f"スキップ: {url} (ページが存在しません)")
            return None

        end_result_tag = soup.find('strong', string='末尾別結果')
        if not end_result_tag:
            print(f"スキップ: {url} (末尾別結果が見つかりません)")
            return None
            
        table = end_result_tag.find_next('table')
        if not table:
            print(f"スキップ: {url} (テーブルが見つかりません)")
            return None

        links = table.find_all('a', href=True)
        for link in links:
            machine_url = link['href']
            if not machine_url.startswith("http"):
                machine_url = url + machine_url

            driver.get(machine_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "a")))

            machine_html = driver.page_source
            machine_soup = BeautifulSoup(machine_html, 'html.parser')
            table2 = machine_soup.find('table', class_='table2')
            if table2:
                rows = table2.find_all('tr')
                data = []
                for row in rows[1:-1]:
                    cols = row.find_all('td')
                    cols = [col.text.strip() for col in cols]
                    if len(cols) == 7 and '1/' in cols[6]:
                        cols[6] = cols[6].replace('1/', '')

                    while len(cols) < 7:
                        cols.append('')

                    data.append(cols)

                all_data.extend(data)

        if not all_data:
            print(f"スキップ: {url} (データが取得できませんでした)")
            return None

        columns = ['dai-num', 'dai-name', 'difference', 'game', 'big', 'reg', 'total']
        df = pd.DataFrame(all_data, columns=columns)
        df.insert(0, 'day', date_formatted)

        def clean_numeric(value):
            value = re.sub(r'[+,]', '', value)
            if '1/' in value:
                return None
            try:
                return pd.to_numeric(value)
            except ValueError:
                return None

        cols_to_clean = ['dai-num', 'difference', 'game', 'big', 'reg', 'total']
        df[cols_to_clean] = df[cols_to_clean].apply(lambda col: col.map(clean_numeric))
        
        # 'big-per' 列と 'reg-per' 列を計算して追加
        df['big-per'] = df['game'] / df['big'] 
        df['reg-per'] = df['game'] / df['reg']
        
        # 'reg' と 'total' 列の間に新しい列を挿入
        cols = df.columns.tolist()
        reg_index = cols.index('reg')
        cols.insert(reg_index + 1, 'big-per')
        cols.insert(reg_index + 2, 'reg-per')
        df = df[cols]
        
        return df
        
    except Exception as e:
        print(f"エラー: {url} の処理中にエラーが発生しました - {str(e)}")
        return None

# スクレイピング開始処理
def start_scraping():
    start_date = get_valid_date()
    if not start_date:
        return

    end_date = datetime.now()
    service = Service(executable_path=r'D:/Users/Documents/python/chromedriver.exe')
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')

    all_dfs = []
    try:
        driver = webdriver.Chrome(service=service, options=options)
        selected_store_name = store_var.get()
        selected_store_url = None
        store_names_urls = fetch_store_names(prefecture_var.get())
        for name, url in store_names_urls:
            if name == selected_store_name:
                selected_store_url = url
                break
        if not selected_store_url:
            messagebox.showerror("エラー", "店名に対応するURLが見つかりませんでした。")
            return
        
        current_date = start_date
        while current_date <= end_date:
            url = generate_url(base_url, selected_store_url, current_date)
            date_formatted = current_date.strftime("%Y-%m-%d")
            print(f"処理中: {date_formatted}")
            df = scrape_page(driver, url, date_formatted)
            if df is not None:
                all_dfs.append(df)
                print(f"成功: {date_formatted} のデータを取得しました")
            current_date += timedelta(days=1)
            time.sleep(1)
    finally:
        driver.quit()

    if all_dfs:
        final_df = pd.concat(all_dfs, ignore_index=True)
        output_path = r"D:\Users\Documents\python\dai-date\final_output.xlsx"
        final_df.to_excel(output_path, index=False)
        messagebox.showinfo("成功", f"データが {output_path} に保存されました。")
    else:
        messagebox.showinfo("情報", "取得できたデータがありませんでした。")

# GUIの作成
root = tk.Tk()
root.title("データスクレイピングツール")

# 都道府県のプルダウンメニュー
tk.Label(root, text="都道府県:").pack(pady=10)
frame = tk.Frame(root)
frame.pack(pady=5)

prefecture_var = tk.StringVar()
prefecture_dropdown = ttk.Combobox(frame, textvariable=prefecture_var)
prefecture_dropdown

# 都道府県のプルダウンメニューの設定
prefecture_var = tk.StringVar()
prefecture_dropdown = ttk.Combobox(frame, textvariable=prefecture_var)
prefecture_dropdown['values'] = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県",
    "岐阜県", "静岡県", "愛知県", "三重県", "滋賀県", "京都府",
    "大阪府", "兵庫県", "奈良県", "和歌山県", "鳥取県", "島根県",
    "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県",
    "高知県", "福岡県", "佐賀県", "長崎県", "熊本県", "大分県",
    "宮崎県", "鹿児島県", "沖縄県"
]  # 全ての都道府県を追加
prefecture_dropdown.grid(row=0, column=0, padx=5)

# 店名の絞り込みボタン
filter_button = tk.Button(frame, text="店名の絞り込み", command=update_store_dropdown)
filter_button.grid(row=0, column=1, padx=5)

# 店名のプルダウンメニュー
tk.Label(root, text="店名:").pack(pady=10)
store_var = tk.StringVar()
store_dropdown = ttk.Combobox(root, textvariable=store_var)
store_dropdown.pack(pady=5)

# 開始日入力
tk.Label(root, text="開始日 (YYYYMMDD形式):").pack(pady=10)
date_entry = tk.Entry(root)
date_entry.pack(pady=5)

# スタートボタン
start_button = tk.Button(root, text="開始", command=start_scraping)
start_button.pack(pady=20)

root.mainloop()
