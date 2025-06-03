# 安裝套件 pip install selenium psutil markupsafe jieba

import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException, NoSuchWindowException, JavascriptException
import time
import threading
import psutil
import subprocess
import re
from collections import Counter
import jieba
import os
import json  # 確保有 import

# =====================
# 主程式：104 求職網頁自動化篩選與高亮工具
# 作者：
# 說明：本程式可自動啟動瀏覽器、滾動頁面、關鍵字/薪資高亮、重複詞標記等，協助求職者高效瀏覽職缺。
# =====================

# =====================
# 全域變數
# =====================
driver = None  # Selenium WebDriver 實例
current_tab = None  # 目前控制的分頁 handle
scrolling = False  # 自動滾動狀態
highlighted_keywords = []  # 已高亮的關鍵字清單

# =====================
# 啟動/連接 Chrome 相關
# =====================
def connect_to_new_browser():
    """
    啟動新的 Chrome 並使用獨立登入資料夾。
    每次都會開啟一個全新、獨立的 Chrome 實例，不會與其他 Chrome 衝突。
    """
    global driver, current_tab
    try:
        # 關閉已有的 driver
        if driver:
            try:
                driver.quit()
            except:
                pass
            driver = None
        # 檢查 chromedriver 是否存在
        chromedriver_path = r"C:\Users\H\Desktop\chromedriver-win64\chromedriver.exe"
        if not os.path.exists(chromedriver_path):
            messagebox.showerror("錯誤", f"未找到 chromedriver: {chromedriver_path}")
            return
        # 檢查 Chrome 瀏覽器是否存在
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            messagebox.showerror("錯誤", f"未找到 Chrome 瀏覽器: {chrome_path}")
            return
        # 設定 Chrome 啟動選項
        chrome_options = Options()
        chrome_options.add_argument(r'--user-data-dir=C:\Chrome104WebBot')  # 使用獨立資料夾，保留登入狀態
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.binary_location = chrome_path
        service = Service(executable_path=chromedriver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        current_tab = driver.current_window_handle
        messagebox.showinfo("提示", "已啟動新的 Chrome，請手動登入 104，之後將自動保留登入狀態。")
        return True
    except Exception as e:
        error_msg = str(e)
        print("無法啟動 Chrome：", error_msg)
        messagebox.showerror("錯誤", f"無法啟動 Chrome：{error_msg}")
        return False

# =====================
# 測試連線
# =====================
def test_connection():
    """
    測試是否可以正常控制瀏覽器。
    若目前分頁不存在，自動切換到現有分頁；若所有分頁都被關閉，顯示錯誤但不重啟瀏覽器。
    """
    global driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        # 嘗試切換到當前分頁
        try:
            if current_tab:
                driver.switch_to.window(current_tab)
        except (NoSuchWindowException, WebDriverException):
            # 當前分頁不存在，嘗試切換到現有分頁
            handles = driver.window_handles
            if handles:
                driver.switch_to.window(handles[0])
                current_tab = handles[0]
            else:
                messagebox.showerror("錯誤", "找不到任何分頁，請手動在瀏覽器開啟新分頁後再點擊測試連接。")
                return
        # 測試執行簡單的 JavaScript
        result = driver.execute_script("return document.title;")
        messagebox.showinfo("連接測試", f"連接正常！當前頁面標題: {result}")
    except Exception as e:
        error_msg = str(e)
        messagebox.showerror("錯誤", f"瀏覽器連接異常: {error_msg}")

# =====================
# 自動滾動功能
# =====================
def start_auto_scroll():
    """
    啟動自動滾動（自動將頁面滾到最底）。
    """
    global scrolling, driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        # 確保在正確的分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        scrolling = True
        def scroll():
            global scrolling
            while scrolling:
                try:
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                except Exception as e:
                    print("捲動錯誤:", e)
                    scrolling = False
                    break
                time.sleep(1)
        t = threading.Thread(target=scroll)
        t.daemon = True
        t.start()
    except Exception as e:
        messagebox.showerror("錯誤", f"開始滾動時出錯: {str(e)}")
        scrolling = False

def stop_auto_scroll():
    """
    停止自動滾動。
    """
    global scrolling
    scrolling = False

# =====================
# 關鍵字高亮功能
# =====================
def highlight_keywords():
    """
    彈出視窗讓使用者輸入關鍵字，並在頁面上高亮顯示。
    """
    global highlighted_keywords, driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        inject_markjs_and_css(driver)
        keywords = simpledialog.askstring("關鍵字", "請輸入關鍵字 (以逗號分隔):")
        if not keywords:
            return
        new_keywords = [keyword.strip() for keyword in keywords.split(",")]
        highlighted_keywords.extend(new_keywords)
        highlighted_keywords = list(set(highlighted_keywords))
        js = f"""
        (function(){{
            var instance = new Mark(document.body);
            // 只移除 keyword-highlight
            document.querySelectorAll('mark.keyword-highlight').forEach(el => {{
                el.outerHTML = el.innerHTML;
            }});
            var keywords = {json.dumps(highlighted_keywords)};
            if(keywords.length === 0) return;
            instance.mark(keywords, {{ className: "keyword-highlight" }});
        }})();
        """
        driver.execute_script(js)
    except Exception as e:
        messagebox.showerror("錯誤", f"高亮關鍵字時出錯: {str(e)}")

def remove_highlight():
    """
    移除所有高亮標記。
    """
    global highlighted_keywords, driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        inject_markjs_and_css(driver)
        js = """
        (function(){
            // 一次移除所有 highlight 標記
            document.querySelectorAll('mark.keyword-highlight, mark.salary-highlight, mark.dup-highlight').forEach(el => {
                el.outerHTML = el.innerHTML;
            });
        })();
        """
        driver.execute_script(js)
        highlighted_keywords = []
    except Exception as e:
        messagebox.showerror("錯誤", f"移除高亮時發生錯誤：{str(e)}")

# =====================
# 薪資高亮功能
# =====================
def salary_highlight():
    """
    彈出視窗讓使用者設定薪資高亮條件，並在頁面上高亮顯示。
    """
    global driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        dialog = tk.Toplevel()
        dialog.title("薪資篩選設定")
        tk.Label(dialog, text="請輸入基準薪資：").grid(row=0, column=0, padx=10, pady=5)
        salary_var = tk.IntVar(value=30000)
        entry_salary = tk.Entry(dialog, textvariable=salary_var, width=10)
        entry_salary.grid(row=0, column=1, padx=10, pady=5)
        check_var1 = tk.BooleanVar(value=True)
        check_var2 = tk.BooleanVar(value=True)
        check_var3 = tk.BooleanVar(value=True)
        tk.Checkbutton(dialog, text="基準薪資以上", variable=check_var1).grid(row=1, column=0, sticky="w", padx=10)
        tk.Checkbutton(dialog, text="基準薪資區間", variable=check_var2).grid(row=2, column=0, sticky="w", padx=10)
        tk.Checkbutton(dialog, text="待遇面議", variable=check_var3).grid(row=3, column=0, sticky="w", padx=10)
        def confirm():
            base_salary = salary_var.get()
            highlight_above = check_var1.get()
            highlight_range = check_var2.get()
            highlight_negotiable = check_var3.get()
            dialog.destroy()
            highlight_salary_on_page(base_salary, highlight_above, highlight_range, highlight_negotiable)
        tk.Button(dialog, text="確定", command=confirm).grid(row=4, column=0, columnspan=2, pady=10)
        dialog.transient(root)
        dialog.grab_set()
        root.wait_window(dialog)
    except Exception as e:
        messagebox.showerror("錯誤", f"薪資高亮設置時發生錯誤：{str(e)}")

def highlight_salary_on_page(base_salary, highlight_above, highlight_range, highlight_negotiable):
    """
    根據設定在頁面上高亮薪資條件。
    """
    global driver, current_tab
    try:
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        inject_markjs_and_css(driver)
        js = f"""
        (function(){{
            if(typeof Mark === 'undefined'){{
                console.error("❌ Mark.js 未載入");
                return;
            }}
            var instance = new Mark(document.body);
            // 只移除 salary-highlight
            document.querySelectorAll('mark.salary-highlight').forEach(el => {{
                el.outerHTML = el.innerHTML;
            }});
            var baseSalary = {base_salary};
            var salaryPattern = /月薪\\s*([\\d,]+)(?:~([\\d,]+))?\\s*元/g;
            var bodyText = document.body.innerHTML;
            var matches = [...bodyText.matchAll(salaryPattern)];
            var highlightTargets = [];
            matches.forEach(match => {{
                var minSalary = parseInt(match[1].replace(/,/g, ''), 10);
                var maxSalary = match[2] ? parseInt(match[2].replace(/,/g, ''), 10) : null;
                if ({'true' if highlight_above else 'false'} && minSalary >= baseSalary) {{
                    highlightTargets.push(match[1]);
                }}
                if ({'true' if highlight_range else 'false'} && minSalary <= baseSalary && maxSalary >= baseSalary) {{
                    highlightTargets.push(match[2]);
                }}
            }});
            if ({'true' if highlight_negotiable else 'false'}) {{
                highlightTargets.push("待遇面議");
            }}
            if (highlightTargets.length > 0) {{
                instance.mark(highlightTargets.filter(Boolean), {{ className: "salary-highlight" }});
            }}
        }})();
        """
        driver.execute_script(js)
    except Exception as e:
        messagebox.showerror("錯誤", f"薪資高亮處理時發生錯誤：{str(e)}")

# =====================
# 分頁切換
# =====================
def switch_to_tab_by_title(title_contains):
    """
    根據分頁標題切換控制分頁。
    """
    global driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    try:
        handles = driver.window_handles
        for handle in handles:
            driver.switch_to.window(handle)
            if title_contains in driver.title:
                current_tab = handle
                messagebox.showinfo("提示", f"已切換到標題包含 '{title_contains}' 的分頁")
                return
        messagebox.showwarning("警告", f"未找到標題包含 '{title_contains}' 的分頁")
    except Exception as e:
        messagebox.showerror("錯誤", f"切換分頁時發生錯誤：{str(e)}")

# =====================
# 注入 Mark.js 與自訂 CSS
# =====================
def inject_markjs_and_css(driver):
    """
    注入 Mark.js 函式庫與自訂高亮 CSS。
    """
    try:
        mark_exists = driver.execute_script("return typeof Mark !== 'undefined';")
        if not mark_exists:
            js_inject = """
            (function(){
                var script = document.createElement('script');
                script.src = "https://cdnjs.cloudflare.com/ajax/libs/mark.js/8.11.1/mark.min.js";
                script.onload = function(){
                    console.log("✅ mark.js 已載入");
                };
                document.head.appendChild(script);
            })();
            """
            driver.execute_script(js_inject)
            for _ in range(10):
                mark_exists = driver.execute_script("return typeof Mark !== 'undefined';")
                if mark_exists:
                    break
                time.sleep(1)
            if not mark_exists:
                print("⚠️ 警告：無法載入 Mark.js 庫")
        css_inject = """
        (function(){
            if(!document.getElementById("custom-highlight-style")){
                var style = document.createElement('style');
                style.id = "custom-highlight-style";
                style.innerHTML = `
                    mark { background: yellow; color: black; }
                    mark.salary-highlight { background: orange; color: black; }
                    mark.dup-highlight { background: yellow; color: black; border-bottom: 2px solid orange; }
                `;
                document.head.appendChild(style);
            }
        })();
        """
        driver.execute_script(css_inject)
    except Exception as e:
        print(f"注入 Mark.js 和 CSS 時發生錯誤：{str(e)}")
        raise

# =====================
# 重複詞標記功能
# =====================
def mark_duplicate_words():
    """
    分析頁面內容，標記出現超過 5 次的重複詞語。
    """
    global driver, current_tab
    try:
        if driver is None:
            messagebox.showerror("錯誤", "請先連接瀏覽器！")
            return
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
        inject_markjs_and_css(driver)
        # 只處理頁面的部分內容（前 20000 字）以減少處理量
        js_get_text = """
        var text = document.body.innerText;
        return text.substring(0, Math.min(text.length, 20000));
        """
        text = driver.execute_script(js_get_text)
        # 先移除所有薪資相關片段
        salary_patterns = [
            r"時薪[\d,]+元",
            r"月薪[\d,]+元(以上)?",
            r"月薪[\d,]+~[\d,]+元",
            r"年薪[\d,]+元(以上)?",
            r"年薪[\d,]+~[\d,]+元"
        ]
        for pat in salary_patterns:
            text = re.sub(pat, "", text)
        # jieba 分詞（限制處理時間最多 5 秒）
        start_time = time.time()
        words = list(jieba.cut(text))
        words = [w.strip() for w in words if w.strip()]
        if time.time() - start_time > 5:
            print("jieba 分詞超時")
        # n-gram 統計（2-4 字）
        ngram_counter = Counter()
        start_time = time.time()
        for n in range(2, 5):
            for i in range(len(text) - n + 1):
                if time.time() - start_time > 5:
                    print("n-gram 統計超時，已處理部分內容")
                    break
                gram = text[i:i+n]
                if '\n' in gram or '\r' in gram or gram.strip() == "":
                    continue
                ngram_counter[gram] += 1
        # 合併 jieba 分詞與 n-gram 統計
        word_count = Counter(words) + ngram_counter
        # 排除清單
        exclude_words = [
            "儲存", "應徵", "上市上櫃", "年終獎金", "節日獎金/禮品", "津貼/補助",
            "分紅配股", "彈性上下班", "小時前處理過履歷", "積極徵才中", "分鐘前聯絡過求職者",
            "人應徵", "優於勞基法特休", "員工旅遊", "停車位", "天內聯絡過求職者",
            "員工餐廳", "部門聚餐", "前往頁面", "健康檢查", "團體保險", "生日假",
            "台北市", "新北市", "桃園市", "新竹市", "苗栗", "台中市",
            "台南市", "高雄市", "補助", "津貼", "獎金", "禮金",
            "有限公司", "處理過", "前聯絡", "積極徵",
        ]
        def is_number_word(word):
            if re.match(r"^[\d,]+$", word):  # 純數字
                return True
            if re.match(r"^[\d,]+元$", word):  # 數字+元
                return True
            if re.match(r"^[\d,]+~[\d,]+$", word):  # 數字~數字
                return True
            return False
        # 找出出現超過 5 次的詞語，且長度大於 1，且不在排除清單，且不是薪資組合
        duplicates = [
            w for w, c in word_count.items()
            if c > 5 and len(w) > 1 and w not in exclude_words and not is_number_word(w)
        ]
        # 依長度排序，最多只取前 500 個
        duplicates = sorted(set(duplicates), key=lambda x: -len(x))[:500]
        if not duplicates:
            messagebox.showinfo("提示", "沒有重複超過 5 次的詞語")
            return
        # 先清除舊的 dup-highlight
        js_clear = """
        (function(){
            document.querySelectorAll('mark.dup-highlight').forEach(el => {
                el.outerHTML = el.innerHTML;
            });
        })();
        """
        driver.execute_script(js_clear)
        # 分批標記，每批 100 個
        batch_size = 100
        for i in range(0, len(duplicates), batch_size):
            batch = duplicates[i:i+batch_size]
            js = f"""
            (function(){{
                var instance = new Mark(document.body);
                var keywords = {batch};
                if(keywords.length === 0) return;
                instance.mark(keywords, {{ className: "dup-highlight", separateWordSearch: false }});
            }})();
            """
            driver.execute_script(js)
            time.sleep(0.5)
        messagebox.showinfo("完成", f"已標記 {len(duplicates)} 個重複詞語")
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        messagebox.showerror("錯誤", f"標記重複詞時發生錯誤：\n{str(e)}")

# =====================
# 關閉瀏覽器
# =====================
def close_browser():
    """
    關閉由本工具啟動的 Chrome 瀏覽器。
    """
    global driver, current_tab
    if driver:
        try:
            driver.quit()
            driver = None
            current_tab = None
            messagebox.showinfo("提示", "已關閉瀏覽器")
        except Exception as e:
            messagebox.showerror("錯誤", f"關閉瀏覽器時發生錯誤：{str(e)}")
    else:
        messagebox.showinfo("提示", "瀏覽器尚未開啟")

# =====================
# GUI 介面
# =====================
root = tk.Tk()
root.title("104 求職網頁工具")

btn_connect_browser = tk.Button(root, text="啟動新的瀏覽器", width=25, command=connect_to_new_browser)
btn_connect_browser.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

btn_test_connection = tk.Button(root, text="測試連接", width=25, command=test_connection)
btn_test_connection.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

btn_start_scroll = tk.Button(root, text="自動end", width=25, command=start_auto_scroll)
btn_start_scroll.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

btn_stop_scroll = tk.Button(root, text="停止end", width=25, command=stop_auto_scroll)
btn_stop_scroll.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

btn_highlight = tk.Button(root, text="highlight 關鍵字", width=25, command=highlight_keywords)
btn_highlight.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

btn_remove_highlight = tk.Button(root, text="移除highlight", width=25, command=remove_highlight)
btn_remove_highlight.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

btn_salary_highlight = tk.Button(root, text="薪資高亮", width=25, command=salary_highlight)
btn_salary_highlight.grid(row=3, column=0, padx=5, pady=5, sticky="ew")

btn_mark_duplicate = tk.Button(root, text="標記重複詞", width=25, command=mark_duplicate_words)
btn_mark_duplicate.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

btn_close_browser = tk.Button(root, text="關閉瀏覽器", width=25, command=close_browser)
btn_close_browser.grid(row=4, column=0, padx=5, pady=5, sticky="ew")

# 讓兩欄都能自動撐滿
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

# 視窗關閉時自動清理資源

def on_closing():
    global driver, scrolling
    scrolling = False
    if driver:
        try:
            driver.quit()
        except:
            pass
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop() 