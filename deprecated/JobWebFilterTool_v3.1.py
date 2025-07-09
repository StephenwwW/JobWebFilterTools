# 104 求職網頁自動化篩選與高亮工具 v3.1 (公司列表高亮整合版)
#
# 功能說明：
# - 自動啟動獨立 Chrome 瀏覽器，避免與日常瀏覽衝突
# - 支援自動滾動、關鍵字高亮、薪資高亮、重複詞標記
# - 適用於 104 求職網頁與一般網頁
# - 具備 GUI 操作介面與「全域」鍵盤快捷鍵
# - 主畫面可直接編輯並儲存各行政區的排除公司名單
#
# 依賴套件：selenium, psutil, markupsafe, jieba, keyboard
# 安裝指令：pip install selenium psutil markupsafe jieba keyboard

import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException
import time
import threading
import os
import json
import keyboard
import jieba
from collections import Counter

# =====================
# 全域變數
# =====================
driver = None
current_tab = None
scrolling = False
highlighted_keywords = []

# =====================
# 多執行緒包裝器
# =====================
def run_threaded(func, *args):
    thread = threading.Thread(target=func, args=args, daemon=True)
    thread.start()

# =====================
# 視窗關閉處理
# =====================
def on_closing(event=None):
    global driver, scrolling
    try:
        print("正在取消全域快捷鍵...")
        keyboard.unhook_all_hotkeys()
    except Exception as e:
        print(f"取消快捷鍵時發生錯誤: {e}")
    scrolling = False
    if driver:
        try:
            driver.quit()
        except:
            pass
    root.destroy()

# =====================
# 核心功能函式
# =====================
def connect_to_new_browser():
    def task():
        global driver, current_tab
        try:
            if driver:
                try: driver.quit()
                except: pass
            globals().update(driver=None, current_tab=None)
            
            chromedriver_path = r"C:\\Users\\H\\Desktop\\chromedriver-win64\\chromedriver.exe"
            if not os.path.exists(chromedriver_path):
                root.after(0, lambda: messagebox.showerror("錯誤", f"未找到 chromedriver: {chromedriver_path}"))
                return
                
            chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            if not os.path.exists(chrome_path):
                root.after(0, lambda: messagebox.showerror("錯誤", f"未找到 Chrome 瀏覽器: {chrome_path}"))
                return
                
            chrome_options = Options()
            chrome_options.add_argument(r'--user-data-dir=C:\\Chrome104WebBot')
            chrome_options.add_argument("--disable-infobars")
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            service = Service(executable_path=chromedriver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {'source': "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
            current_tab = driver.current_window_handle
            root.after(0, lambda: messagebox.showinfo("提示", "已啟動新的 Chrome。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"無法啟動 Chrome：{e}"))
    run_threaded(task)

def test_connection():
    if driver is None: messagebox.showerror("錯誤", "請先連接瀏覽器！"); return
    def task():
        try:
            result = driver.execute_script("return document.title;")
            root.after(0, lambda: messagebox.showinfo("連接測試", f"連接正常！當前頁面標題: {result}"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"瀏覽器連接異常: {e}"))
    run_threaded(task)

def start_auto_scroll():
    global scrolling
    if driver is None: messagebox.showerror("錯誤", "請先連接瀏覽器！"); return
    if scrolling: return
    scrolling = True
    def scroll_task():
        global scrolling
        while scrolling:
            try:
                if driver: driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                else: break
                time.sleep(1)
            except Exception:
                scrolling = False
                break
    run_threaded(scroll_task)

def stop_auto_scroll():
    global scrolling
    scrolling = False

def inject_markjs_and_css(driver_instance):
    try:
        mark_exists = driver_instance.execute_script("return typeof Mark !== 'undefined';")
        if not mark_exists:
            js_inject = """
            (function(){
                var script = document.createElement('script');
                script.src = "https://cdnjs.cloudflare.com/ajax/libs/mark.js/8.11.1/mark.min.js";
                script.onload = function(){ console.log("mark.js 已載入"); };
                document.head.appendChild(script);
            })();
            """
            driver_instance.execute_script(js_inject)
            time.sleep(1.5)
        css_inject = """
        (function(){
            if(!document.getElementById("custom-highlight-style")){
                var style = document.createElement('style');
                style.id = "custom-highlight-style";
                style.innerHTML = `
                    mark { background: yellow !important; border-radius: 2px; }
                    mark.salary-highlight { background: orange !important; border-radius: 2px; }
                    mark.dup-highlight { background: yellow !important; border-bottom: 2px solid orange; border-radius: 2px; }
                `;
                document.head.appendChild(style);
            }
        })();
        """
        driver_instance.execute_script(css_inject)
    except Exception as e:
        print(f"注入 Mark.js 和 CSS 時發生錯誤：{e}")

# ==============================================================================
# 【主要修改處】
# ==============================================================================
def highlight_keywords():
    if driver is None: 
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return

    # 步驟 1: 在主執行緒中，從 GUI 獲取所有需要高亮的關鍵字
    
    # 從彈出視窗獲取額外關鍵字
    user_input_str = simpledialog.askstring("高亮關鍵字", "請輸入額外關鍵字 (以逗號分隔)，\n將與右側公司列表一同高亮顯示。")
    user_keywords = [k.strip() for k in user_input_str.split(',') if k.strip()] if user_input_str else []

    # 從右側文字區塊獲取公司名單
    try:
        company_list_text = result_text.get("1.0", tk.END).strip()
        lines = company_list_text.split('\n')
        # 確保文字區塊格式正確，並從第三行開始讀取
        if len(lines) > 2: 
            company_names = [line.strip() for line in lines[2:] if line.strip()]
        else:
            company_names = []
    except Exception as e:
        print(f"從文字區塊讀取公司列表時發生錯誤: {e}")
        company_names = []

    # 合併兩個來源的關鍵字
    final_keywords_to_add = user_keywords + company_names
    
    if not final_keywords_to_add:
        messagebox.showinfo("提示", "沒有提供任何關鍵字或公司名稱可供高亮。")
        return

    # 步驟 2: 將耗時的網頁高亮工作，連同關鍵字列表，交給背景執行緒處理
    def do_highlighting_task(keywords):
        global highlighted_keywords
        try:
            # 注入高亮腳本
            inject_markjs_and_css(driver)

            # 將新關鍵字加入全域列表，並移除重複項
            highlighted_keywords.extend(keywords)
            highlighted_keywords = sorted(list(set(highlighted_keywords)), key=len, reverse=True)

            # 準備並執行 JavaScript
            js = f"""(function() {{
                var instance = new Mark(document.body);
                // 使用更新後的完整列表進行高亮
                var keywords_to_mark = {json.dumps(highlighted_keywords)};
                // 先清除舊的同類型高亮，再上新的，避免疊加
                instance.unmark({{ className: 'keyword-highlight', done: function() {{
                    if(keywords_to_mark.length > 0) {{
                        instance.mark(keywords_to_mark, {{ "className": "keyword-highlight" }});
                    }}
                }} }});
            }})();"""
            driver.execute_script(js)
            root.after(0, lambda: messagebox.showinfo("完成", f"已高亮 {len(keywords)} 個新的關鍵字/公司。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"高亮關鍵字時出錯: {e}"))

    # 啟動背景執行緒
    run_threaded(do_highlighting_task, final_keywords_to_add)

def remove_highlight():
    if driver is None: messagebox.showerror("錯誤", "請先連接瀏覽器！"); return
    def do_unhighlight():
        try:
            inject_markjs_and_css(driver)
            js = """(function(){
                new Mark(document.body).unmark({ className: 'keyword-highlight' });
                new Mark(document.body).unmark({ className: 'salary-highlight' });
                new Mark(document.body).unmark({ className: 'dup-highlight' });
            })();"""
            driver.execute_script(js)
            globals()['highlighted_keywords'] = []
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"移除高亮時發生錯誤：{e}"))
    run_threaded(do_unhighlight)

def highlight_salary_on_page(base_salary, highlight_above, highlight_range, highlight_negotiable):
    try:
        inject_markjs_and_css(driver)
        js = f"""
        (function(){{
            if(typeof Mark === 'undefined') {{ console.error("Mark.js not loaded"); return; }}
            var instance = new Mark(document.body);
            instance.unmark({{ className: 'salary-highlight', done: function() {{
                var highlightTargets = [];
                var bodyText = document.body.innerText;
                var patterns = [
                    /月薪\\s*([\\d,]+)\\s*~\\s*([\\d,]+)\\s*元/g, /月薪\\s*([\\d,]+)\\s*元/g,
                    /年薪\\s*([\\d,]+)\\s*~\\s*([\\d,]+)\\s*元/g, /年薪\\s*([\\d,]+)\\s*元/g
                ];
                patterns.forEach(function(pattern) {{
                    var match;
                    while ((match = pattern.exec(bodyText)) !== null) {{
                        var isYearly = match[0].includes('年薪');
                        var minSalary = parseInt(match[1].replace(/,/g, ''), 10);
                        var maxSalary = match[2] ? parseInt(match[2].replace(/,/g, ''), 10) : minSalary;
                        var monthlyMin = isYearly ? Math.round(minSalary / 12) : minSalary;
                        var monthlyMax = isYearly ? Math.round(maxSalary / 12) : maxSalary;
                        if (({'true' if highlight_above else 'false'} && monthlyMin >= {base_salary}) || 
                            ({'true' if highlight_range else 'false'} && {base_salary} >= monthlyMin && {base_salary} <= monthlyMax)) {{
                            highlightTargets.push(match[0]);
                        }}
                    }}
                }});
                if ({'true' if highlight_negotiable else 'false'}) {{
                    highlightTargets.push("待遇面議");
                }}
                if (highlightTargets.length > 0) {{
                    instance.mark(highlightTargets, {{ "className": "salary-highlight", "separateWordSearch": false }});
                }}
            }} }});
        }})();
        """
        driver.execute_script(js)
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("錯誤", f"薪資高亮處理時發生錯誤：{e}"))

def salary_highlight():
    if driver is None: messagebox.showerror("錯誤", "請先連接瀏覽器！"); return
    dialog = tk.Toplevel(root)
    dialog.title("薪資篩選設定")
    tk.Label(dialog, text="請輸入月薪基準：").grid(row=0, column=0, padx=10, pady=5)
    salary_var = tk.IntVar(value=40000)
    tk.Entry(dialog, textvariable=salary_var, width=10).grid(row=0, column=1, padx=10, pady=5)
    check_var1 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="高於基準", variable=check_var1).grid(row=1, column=0, sticky="w", padx=10)
    check_var2 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="包含基準", variable=check_var2).grid(row=2, column=0, sticky="w", padx=10)
    check_var3 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="面議/其他", variable=check_var3).grid(row=3, column=0, sticky="w", padx=10)
    def confirm():
        base_salary, highlight_above, highlight_range, highlight_negotiable = salary_var.get(), check_var1.get(), check_var2.get(), check_var3.get()
        dialog.destroy()
        run_threaded(highlight_salary_on_page, base_salary, highlight_above, highlight_range, highlight_negotiable)
    tk.Button(dialog, text="確定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)
    dialog.transient(root); dialog.grab_set(); root.wait_window(dialog)

def mark_duplicate_words():
    if driver is None: messagebox.showerror("錯誤", "請先連接瀏覽器！"); return
    def task():
        try:
            inject_markjs_and_css(driver)
            text = driver.execute_script("return document.body.innerText.substring(0, 30000);")
            exclude_words = { "儲存", "應徵", "上市上櫃", "年終獎金", "相關"} # etc.
            words = jieba.lcut(text)
            word_counts = Counter(w for w in words if len(w) > 1 and w not in exclude_words and not w.isdigit())
            duplicates = [w for w, c in word_counts.items() if c > 5]
            duplicates = sorted(list(set(duplicates)), key=len, reverse=True)[:200]
            if not duplicates:
                root.after(0, lambda: messagebox.showinfo("提示", "沒有找到重複5次以上的詞語。"))
                return
            js_highlight = f"""
            (function(){{
                var instance = new Mark(document.body);
                var keywords = {json.dumps(duplicates)};
                instance.unmark({{ className: 'dup-highlight', done: function(){{
                    instance.mark(keywords, {{ className: "dup-highlight", separateWordSearch: false }});
                }} }});
            }})();
            """
            driver.execute_script(js_highlight)
            root.after(0, lambda: messagebox.showinfo("完成", f"已標記 {len(duplicates)} 個重複詞語。"))
        except Exception as e:
            root.after(0, lambda: messagebox.showerror("錯誤", f"標記重複詞時發生錯誤：\n{e}"))
    run_threaded(task)

def close_browser():
    def task():
        if driver:
            try:
                driver.quit()
                globals().update(driver=None, current_tab=None, scrolling=False)
                root.after(0, lambda: messagebox.showinfo("提示", "已關閉瀏覽器"))
            except Exception as e: 
                root.after(0, lambda: messagebox.showerror("錯誤", f"關閉瀏覽器時發生錯誤：{e}"))
        else:
            root.after(0, lambda: messagebox.showinfo("提示", "瀏覽器尚未開啟"))
    run_threaded(task)

def save_company_filters():
    city, district = city_var.get(), district_var.get()
    if not city or not district: 
        messagebox.showwarning("警告", "請先選擇城市和行政區。")
        return
        
    full_text = result_text.get("1.0", tk.END).strip()
    company_names = [line.strip() for line in full_text.split('\n')[2:] if line.strip()]
    
    company_filters_path = os.path.join(os.path.dirname(__file__), "company_filters.json")
    try:
        with open(company_filters_path, "r", encoding="utf-8") as f:
            company_data = json.load(f)
            
        if city not in company_data:
            company_data[city] = {}
        company_data[city][district] = company_names
        
        with open(company_filters_path, "w", encoding="utf-8") as f:
            json.dump(company_data, f, ensure_ascii=False, indent=4)
            
        messagebox.showinfo("成功", f"已儲存 {city} {district} 的排除公司名單。")
    except FileNotFoundError:
        messagebox.showerror("錯誤", f"找不到公司設定檔：{company_filters_path}")
    except Exception as e:
        messagebox.showerror("錯誤", f"儲存失敗：{e}")

def register_global_hotkeys():
    print("正在註冊全域快捷鍵...")
    def create_callback(func):
        def schedule_on_main_thread():
            # print(f"[{time.ctime()}] 快捷鍵觸發: '{func.__name__}'")
            # 為了避免洗版，暫時關閉日誌
            root.after(0, func)
        return schedule_on_main_thread
    hotkey_map = {
        'alt+a': connect_to_new_browser, 'alt+w': test_connection,
        'alt+s': start_auto_scroll, 'alt+x': stop_auto_scroll,
        'alt+d': highlight_keywords, 'alt+r': remove_highlight,
        'alt+e': salary_highlight, 'alt+f': mark_duplicate_words,
        'alt+q': close_browser, 'alt+z': save_company_filters
    }
    for key, func in hotkey_map.items():
        keyboard.add_hotkey(key, create_callback(func))
    print("全域快捷鍵已註冊成功。")
    print("="*40)

# =====================
# GUI 介面
# =====================
root = tk.Tk()
root.title("104 求職網頁自動化篩選與高亮工具 v3.1")

btn_frame = tk.Frame(root)
btn_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")
buttons = {
    "啟動/連接 (Alt+A)": connect_to_new_browser, "測試連接 (Alt+W)": test_connection,
    "開始滾動 (Alt+S)": start_auto_scroll, "停止滾動 (Alt+X)": stop_auto_scroll,
    "高亮關鍵字 (Alt+D)": highlight_keywords, "移除高亮 (Alt+R)": remove_highlight,
    "薪資高亮 (Alt+E)": salary_highlight, "標記重複詞 (Alt+F)": mark_duplicate_words,
    "關閉瀏覽器 (Alt+Q)": close_browser
}
row, col = 0, 0
for text, command in buttons.items():
    tk.Button(btn_frame, text=text, width=25, command=command).grid(row=row, column=col, padx=5, pady=5, sticky="ew")
    col += 1
    if col > 1: col = 0; row += 1

right_frame = tk.Frame(root)
right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")
with open(os.path.join(os.path.dirname(__file__), "taiwan_districts.json"), "r", encoding="utf-8") as f:
    districts_data = json.load(f)
city_list = list(districts_data.keys())
with open(os.path.join(os.path.dirname(__file__), "company_filters.json"), "r", encoding="utf-8") as f:
    company_filters = json.load(f)
city_var = tk.StringVar()
district_var = tk.StringVar()
tk.Label(right_frame, text="城市").grid(row=0, column=0, sticky="w", padx=5, pady=5)
city_menu = tk.OptionMenu(right_frame, city_var, *city_list)
city_menu.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
tk.Label(right_frame, text="行政區").grid(row=1, column=0, sticky="w", padx=5, pady=5)
district_menu = tk.OptionMenu(right_frame, district_var, "")
district_menu.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
result_text = tk.Text(right_frame, height=23, width=60, font=("Consolas", 11))
result_text.grid(row=0, column=2, rowspan=19, padx=10, pady=5, sticky="n")
result_text.insert("end", "請選擇城市與行政區\n"); result_text.config(state="disabled")

def update_districts(*args):
    city = city_var.get()
    districts = list(districts_data.get(city, {}).get("districts", {}).keys())
    menu = district_menu['menu']
    menu.delete(0, 'end')
    if districts:
        for d in districts: menu.add_command(label=d, command=tk._setit(district_var, d))
        district_var.set(districts[0])
    else: district_var.set("")

def on_city_or_district_change(*args):
    with open(os.path.join(os.path.dirname(__file__), "company_filters.json"), "r", encoding="utf-8") as f:
        current_company_filters = json.load(f)
    city, district = city_var.get(), district_var.get()
    result_text.config(state="normal"); result_text.delete("1.0", "end")
    result_text.insert("end", f"[行政區] {city} {district}\n[排除公司]\n")
    if city in current_company_filters and district in current_company_filters.get(city, {}):
        for cname in current_company_filters[city][district]: result_text.insert("end", f"{cname}\n")

city_var.trace_add('write', update_districts)
city_var.trace_add('write', on_city_or_district_change)
district_var.trace_add('write', on_city_or_district_change)
btn_save_companies = tk.Button(right_frame, text="儲存排除公司名單 (Alt+Z)", command=save_company_filters)
btn_save_companies.grid(row=20, column=2, padx=10, pady=5, sticky="ew")
if city_list: city_var.set(city_list[0])

# =====================
# 程式啟動點
# =====================
try:
    register_global_hotkeys()
except Exception as e:
    if "permission" in str(e).lower() or ("ImportError" in str(e) and "root" in str(e)):
         messagebox.showwarning("權限錯誤", "無法註冊全域快捷鍵。\n請嘗試以「系統管理員身分」執行此程式。")
    else:
        messagebox.showerror("快捷鍵錯誤", f"註冊全域快捷鍵時發生錯誤: {e}")

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()