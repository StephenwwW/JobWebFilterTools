# 104 求職網頁自動化篩選與高亮工具 v4.7 (薪資高亮修復整合版)
#
# 功能說明：
# - [修復] 根據使用者提供的 v3.1 有效版本，將薪資高亮的核心邏輯還原，解決新版失效問題。
# - [整合] 將 v3.1 的薪資高亮邏輯與 v4.x 的現代化腳本框架（Shadow DOM、重複職缺高亮等）結合。
# - 保留所有 v4.x 版本的功能優勢，同時確保核心功能的穩定性。
#
# 依賴套件：selenium, psutil, markupsafe, jieba, keyboard
# 安裝指令：pip install --upgrade selenium psutil markupsafe jieba keyboard

import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, NoSuchWindowException
import time
import threading
import os
import json
import keyboard
import jieba
from collections import Counter
import random

# =====================
# 全域變數
# =====================
driver = None
scrolling = False
MARK_JS_CONTENT = None # 用來快取 mark.js 的內容

# =====================
# 多執行緒包裝器
# =====================
def run_threaded(func, *args, **kwargs):
    thread = threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True)
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
# 核心功能函式 (v4.7 修改)
# =====================
def switch_to_104_tab(show_success_message=False):
    if driver is None:
        root.after(0, lambda: messagebox.showerror("錯誤", "請先啟動/連接瀏覽器！"))
        return False
    try:
        all_tabs = driver.window_handles
        for tab in all_tabs:
            driver.switch_to.window(tab)
            if "104.com.tw" in driver.current_url:
                if show_success_message:
                     current_url = driver.current_url
                     root.after(0, lambda: messagebox.showinfo("成功", f"已成功鎖定 104.com.tw 分頁。\n目前網址: {current_url}"))
                return True
        root.after(0, lambda: messagebox.showwarning("未找到", "未找到開啟的 104.com.tw 相關分頁。\n請先手動開啟 104 求職網頁。"))
        return False
    except (NoSuchWindowException, WebDriverException) as e:
        globals().update(driver=None, scrolling=False)
        root.after(0, lambda err=e: messagebox.showerror("瀏覽器錯誤", f"瀏覽器視窗已關閉或無回應。\n請重新啟動/連接瀏覽器。\n\n錯誤訊息: {err}"))
        return False
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("未知錯誤", f"切換分頁時發生未知錯誤: {err}"))
        return False

def connect_to_new_browser():
    def task():
        global driver
        try:
            if driver:
                try: driver.quit()
                except: pass
            globals().update(driver=None)
            
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
            
            user_agents = [
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            ]
            chrome_options.add_argument(f'user-agent={random.choice(user_agents)}')
            
            service = Service()
            driver = webdriver.Chrome(service=service, options=chrome_options)
            
            driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {'source': "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"})
            root.after(0, lambda: messagebox.showinfo("提示", "已啟動新的 Chrome。請手動開啟或導向 104.com.tw 網頁後再進行操作。"))
        except Exception as e:
            root.after(0, lambda err=e: messagebox.showerror("錯誤", f"無法啟動 Chrome：{err}\n\n請確認：\n1. Selenium已是最新版\n2. 網路連線正常"))
    run_threaded(task)

def focus_and_test_connection():
    run_threaded(switch_to_104_tab, show_success_message=True)

def start_auto_scroll():
    if not switch_to_104_tab(): return
    js_scroll = """
        window.scrollInterval = setInterval(function() {
            window.scrollTo(0, document.body.scrollHeight);
        }, 1000);
        window.stopScrolling = function() {
            clearInterval(window.scrollInterval);
        };
    """
    run_threaded(driver.execute_script, js_scroll)

def stop_auto_scroll():
    if not switch_to_104_tab(): return
    run_threaded(driver.execute_script, "if(typeof window.stopScrolling === 'function') { window.stopScrolling(); }")

def execute_recursive_script(script_name, *args):
    if not switch_to_104_tab():
        return
        
    global MARK_JS_CONTENT
    if script_name in ['highlight', 'unmark', 'highlightSalary'] and MARK_JS_CONTENT is None:
        try:
            script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'mark.min.js')
            with open(script_path, 'r', encoding='utf-8') as f:
                MARK_JS_CONTENT = f.read()
        except FileNotFoundError:
            messagebox.showerror("錯誤", "找不到 'mark.min.js' 檔案。\n請下載它並與本程式放在同一個資料夾。")
            return

    json_args = ', '.join(json.dumps(arg) for arg in args)
    full_js = f"""
    (async function() {{
        const styleId = 'custom-highlight-style';
        if (!document.getElementById(styleId)) {{
            const style = document.createElement('style');
            style.id = styleId;
            style.innerHTML = `
                mark.keyword-highlight {{ background: yellow !important; color: black !important; }}
                mark.salary-highlight {{ background: orange !important; color: black !important; }}
                .duplicate-job-highlight {{
                    background-color: rgba(211, 211, 211, 0.6) !important;
                    border: 2px dashed #999 !important;
                }}
                .duplicate-job-highlight:hover::before {{
                    content: ' 重複職缺 '; position: absolute; top: 0; left: 0;
                    background-color: #dc3545; color: white; padding: 2px 5px;
                    font-size: 12px; border-radius: 3px; z-index: 100;
                }}
            `;
            document.head.appendChild(style);
        }}

        async function recursiveExecute(doc, action) {{
            let count = 0;
            try {{
                const result = await action(doc);
                if (typeof result === 'number') count += result;
            }} catch(e) {{ console.warn('Error executing action in a document context:', e); }}

            for (const el of doc.querySelectorAll('*')) {{
                if (el.shadowRoot) count += await recursiveExecute(el.shadowRoot, action);
            }}
            return count;
        }}

        const actions = {{
            highlight: (keywords) => (context) => new Mark(context).mark(keywords, {{ "className": "keyword-highlight" }}),
            unmark: () => (context) => new Mark(context).unmark(),
            // [v4.7 修復] 採用 v3.1 驗證有效的薪資高亮邏輯
            highlightSalary: (options) => (context) => {{
                return new Promise((resolve) => {{
                    var instance = new Mark(context);
                    instance.unmark({{ className: 'salary-highlight', done: function() {{
                        var highlightTargets = [];
                        // 關鍵改動：直接讀取整個 document.body 的文字，而不是 context 的。
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
                                if ((options.highlight_above && monthlyMin >= options.base_salary) || 
                                    (options.highlight_range && options.base_salary >= monthlyMin && options.base_salary <= monthlyMax)) {{
                                    highlightTargets.push(match[0]);
                                }}
                            }}
                        }});
                        if (options.highlight_negotiable) {{
                            highlightTargets.push("待遇面議");
                        }}
                        if (highlightTargets.length > 0) {{
                            // 高亮操作依然在 context (document 或 shadowRoot) 中執行
                            instance.mark(highlightTargets, {{ "className": "salary-highlight", "separateWordSearch": false }});
                        }}
                        resolve(highlightTargets.length);
                    }} }});
                }});
            }},
            highlightDuplicates: (options) => async (context) => {{
                let duplicateCount = 0;
                const seenJobs = new Set();
                const jobCards = context.querySelectorAll(options.cardSelector);
                jobCards.forEach(card => {{
                    const companyEl = card.querySelector(options.companySelector);
                    const jobEl = card.querySelector(options.jobSelector);
                    if (companyEl && jobEl) {{
                        const key = (companyEl.innerText || "").trim() + '|' + (jobEl.innerText || "").trim();
                        if (key !== '|') {{
                            if (seenJobs.has(key)) {{
                                card.classList.add('duplicate-job-highlight');
                                duplicateCount++;
                            }} else {{
                                seenJobs.add(key);
                                card.classList.remove('duplicate-job-highlight');
                            }}
                        }}
                    }}
                }});
                return duplicateCount;
            }},
            unmarkDuplicates: (options) => async (context) => {{
                 const cards = context.querySelectorAll(options.cardSelector);
                 cards.forEach(card => card.classList.remove('duplicate-job-highlight'));
                 return cards.length;
            }}
        }};

        try {{
            if (typeof Mark === 'undefined' && "{script_name}".includes("highlight")) {{
                {MARK_JS_CONTENT or 'console.error("mark.js not loaded")'}
            }}
            const totalCount = await recursiveExecute(document, actions['{script_name}']({json_args}));
            return {{ success: true, count: totalCount }};
        }} catch(e) {{
            console.error('執行遞迴腳本時出錯:', e);
            return {{ success: false, error: e.toString() }};
        }}
    }})();
    """
    
    def callback(result):
        if result and result.get('success'):
            count = result.get('count', 0)
            message = ""
            if script_name == 'highlightDuplicates':
                message = f"掃描完畢，共標記了 {count} 個重複職缺。"
            elif script_name == 'unmarkDuplicates':
                message = "已清除所有重複職缺高亮。"
            elif script_name == 'highlightSalary':
                message = "薪資高亮請求已送出。"
            
            if message:
                root.after(0, lambda: messagebox.showinfo("完成", message))

    run_threaded(lambda: callback(driver.execute_script(full_js)))

def highlight_keywords():
    user_input_str = simpledialog.askstring("高亮關鍵字", "請輸入額外關鍵字 (以逗號分隔)，\n將與右側公司列表一同高亮顯示。")
    if not user_input_str: return
    user_keywords = [k.strip() for k in user_input_str.split(',') if k.strip()]
    try:
        company_list_text = result_text.get("1.0", tk.END).strip()
        lines = company_list_text.split('\n')
        company_names = [line.strip() for line in lines[2:] if line.strip()] if len(lines) > 2 else []
    except Exception as e:
        print(f"從文字區塊讀取公司列表時發生錯誤: {e}"); company_names = []
    final_keywords = list(set(user_keywords + company_names))
    if not final_keywords: messagebox.showinfo("提示", "沒有提供任何關鍵字。"); return
    
    execute_recursive_script('highlight', final_keywords)
    messagebox.showinfo("完成", f"已送出 {len(final_keywords)} 個關鍵字的高亮請求。")

def remove_highlight():
    execute_recursive_script('unmark')
    messagebox.showinfo("完成", "已送出清除高亮的請求。")

def salary_highlight():
    dialog = tk.Toplevel(root)
    dialog.title("薪資篩選設定")
    tk.Label(dialog, text="請輸入月薪基準：").grid(row=0, column=0, padx=10, pady=5)
    salary_var = tk.IntVar(value=40000)
    tk.Entry(dialog, textvariable=salary_var, width=10).grid(row=0, column=1, padx=10, pady=5)
    check_var1 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="高於基準", variable=check_var1).grid(row=1, column=0, sticky="w", padx=10)
    check_var2 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="包含基準", variable=check_var2).grid(row=2, column=0, sticky="w", padx=10)
    check_var3 = tk.BooleanVar(value=True); tk.Checkbutton(dialog, text="面議/其他", variable=check_var3).grid(row=3, column=0, sticky="w", padx=10)
    def confirm():
        options = {
            "base_salary": salary_var.get(),
            "highlight_above": check_var1.get(),
            "highlight_range": check_var2.get(),
            "highlight_negotiable": check_var3.get()
        }
        dialog.destroy()
        execute_recursive_script('highlightSalary', options)
    tk.Button(dialog, text="確定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)
    dialog.transient(root); dialog.grab_set(); root.wait_window(dialog)

def highlight_duplicate_jobs():
    options = {
        "cardSelector": "article.b-block--top-bord",
        "companySelector": "a[data-qa-id='companyName']",
        "jobSelector": "a[data-qa-id='jobName']"
    }
    execute_recursive_script('highlightDuplicates', options)

def remove_duplicate_highlight():
    options = { "cardSelector": "article.b-block--top-bord" }
    execute_recursive_script('unmarkDuplicates', options)

def close_browser():
    def task():
        global driver
        if driver:
            try:
                driver.quit()
                globals().update(driver=None, scrolling=False)
                root.after(0, lambda: messagebox.showinfo("提示", "已關閉瀏覽器"))
            except Exception as e:
                root.after(0, lambda err=e: messagebox.showerror("錯誤", f"關閉瀏覽器時發生錯誤：{err}"))
        else:
            root.after(0, lambda: messagebox.showinfo("提示", "瀏覽器尚未開啟"))
    run_threaded(task)

def save_company_filters():
    city, district = city_var.get(), district_var.get()
    if not city or not district: messagebox.showwarning("警告", "請先選擇城市和行政區。"); return
    full_text = result_text.get("1.0", tk.END).strip()
    company_names = [line.strip() for line in full_text.split('\n')[2:] if line.strip()]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    company_filters_path = os.path.join(script_dir, "company_filters.json")
    try:
        if not os.path.exists(company_filters_path):
            with open(company_filters_path, "w", encoding="utf-8") as f: json.dump({}, f)
        with open(company_filters_path, "r", encoding="utf-8") as f: company_data = json.load(f)
        if city not in company_data: company_data[city] = {}
        company_data[city][district] = company_names
        with open(company_filters_path, "w", encoding="utf-8") as f: json.dump(company_data, f, ensure_ascii=False, indent=4)
        messagebox.showinfo("成功", f"已儲存 {city} {district} 的排除公司名單。")
    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("錯誤", f"儲存失敗：{err}"))

def register_global_hotkeys():
    print("正在註冊全域快捷鍵...")
    def create_callback(func):
        def schedule_on_main_thread(): root.after(0, func)
        return schedule_on_main_thread
        
    hotkey_map = {
        'alt+a': connect_to_new_browser, 'alt+w': focus_and_test_connection,
        'alt+s': start_auto_scroll, 'alt+x': stop_auto_scroll,
        'alt+d': highlight_keywords, 'alt+r': remove_highlight,
        'alt+e': salary_highlight, 'alt+q': close_browser, 
        'alt+z': save_company_filters,
        'alt+f': lambda: switch_to_104_tab(show_success_message=True),
        'alt+c': highlight_duplicate_jobs, 'alt+v': remove_duplicate_highlight
    }
    for key, func in hotkey_map.items():
        try:
            keyboard.add_hotkey(key, create_callback(func))
        except Exception as e: print(f"註冊快捷鍵 {key} 失敗: {e}")
    print("全域快捷鍵註冊完畢。"); print("="*40)

# =====================
# GUI 介面
# =====================
root = tk.Tk()
root.title("求職網頁自動化篩選工具 v4.7")

# ... (其餘 GUI 介面程式碼維持不變，此處為節省篇幅省略) ...
# ... 您可以從 v4.5/v4.6 複製完整的 GUI 區塊 ...
btn_frame = tk.Frame(root)
btn_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")
buttons = {
    "啟動/連接 (Alt+A)": connect_to_new_browser, 
    "測試/重連104分頁 (Alt+W)": focus_and_test_connection,
    "鎖定104分頁 (Alt+F)": lambda: switch_to_104_tab(show_success_message=True),
    "開始滾動 (Alt+S)": start_auto_scroll, 
    "停止滾動 (Alt+X)": stop_auto_scroll,
    "高亮關鍵字 (Alt+D)": highlight_keywords, 
    "移除高亮 (Alt+R)": remove_highlight,
    "高亮重複職缺 (Alt+C)": highlight_duplicate_jobs,
    "移除重複高亮 (Alt+V)": remove_duplicate_highlight,
    "薪資高亮 (Alt+E)": salary_highlight, 
    "關閉瀏覽器 (Alt+Q)": close_browser,
}
row, col = 0, 0
for text, command in buttons.items():
    tk.Button(btn_frame, text=text, width=25, command=command).grid(row=row, column=col, padx=5, pady=5, sticky="ew")
    col += 1;
    if col > 1: col = 0; row += 1
right_frame = tk.Frame(root)
right_frame.grid(row=0, column=1, padx=10, pady=10, sticky="n")
script_dir = os.path.dirname(os.path.abspath(__file__))
districts_data_path = os.path.join(script_dir, "taiwan_districts.json")
company_filters_path = os.path.join(script_dir, "company_filters.json")
try:
    with open(districts_data_path, "r", encoding="utf-8") as f: districts_data = json.load(f)
    city_list = list(districts_data.keys())
except FileNotFoundError:
    messagebox.showerror("錯誤", f"找不到地區設定檔: {districts_data_path}")
    districts_data, city_list = {}, []
if not os.path.exists(company_filters_path):
    with open(company_filters_path, 'w', encoding='utf-8') as f: json.dump({}, f)
try:
    with open(company_filters_path, "r", encoding="utf-8") as f: company_filters = json.load(f)
except (json.JSONDecodeError, FileNotFoundError) as e:
    messagebox.showerror("錯誤", f"讀取公司設定檔失敗: {e}"); company_filters = {}
city_var = tk.StringVar()
district_var = tk.StringVar()
tk.Label(right_frame, text="城市").grid(row=0, column=0, sticky="w", padx=5, pady=5)
city_menu = tk.OptionMenu(right_frame, city_var, *city_list if city_list else [""])
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
    try:
        with open(company_filters_path, "r", encoding="utf-8") as f: current_company_filters = json.load(f)
    except Exception: current_company_filters = {}
    city, district = city_var.get(), district_var.get()
    if not city or not district: return
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