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

# 全局變數
driver = None
current_tab = None
scrolling = False
highlighted_keywords = []

def connect_to_new_browser():
    """啟動新的 Chrome 並使用獨立登入資料夾"""
    global driver, current_tab
    try:
        # 关闭已有的driver
        if driver:
            try:
                driver.quit()
            except:
                pass
            driver = None
            
        # 检查chromedriver是否存在
        chromedriver_path = r"C:\Users\H\Desktop\chromedriver-win64\chromedriver.exe"
        if not os.path.exists(chromedriver_path):
            messagebox.showerror("错误", f"未找到chromedriver: {chromedriver_path}")
            return
            
        # 检查Chrome浏览器是否存在
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            messagebox.showerror("错误", f"未找到Chrome浏览器: {chrome_path}")
            return
            
        chrome_options = Options()
        chrome_options.add_argument(r'--user-data-dir=C:\Chrome104WebBot')
        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.binary_location = chrome_path

        # 指定 chromedriver 完整路徑
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

def test_connection():
    """测试是否可以正常控制浏览器"""
    global driver, current_tab
    if driver is None:
        messagebox.showerror("错误", "请先连接浏览器！")
        return
        
    try:
        # 確保已切換到當前分頁
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
                
        # 测试执行简单的JavaScript
        result = driver.execute_script("return document.title;")
        messagebox.showinfo("连接测试", f"连接正常！当前页面标题: {result}")
    except Exception as e:
        error_msg = str(e)
        messagebox.showerror("错误", f"浏览器连接异常: {error_msg}")
        # 尝试重新连接
        if "session" in error_msg.lower() or "connection" in error_msg.lower():
            if messagebox.askyesno("重新连接", "连接已断开，是否尝试重新连接?"):
                connect_to_new_browser()

def start_auto_scroll():
    global scrolling, driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
        
    try:
        # 确保在正确的标签页
        if current_tab:
            try:
                driver.switch_to.window(current_tab)
            except NoSuchWindowException:
                current_tab = driver.current_window_handle
                
        scrolling = True
        def scroll():
            global scrolling  # 在线程函数中也声明scrolling为全局变量
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
        messagebox.showerror("错误", f"开始滚动时出错: {str(e)}")
        scrolling = False

def stop_auto_scroll():
    global scrolling
    scrolling = False

def highlight_keywords():
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
            var keywords = {highlighted_keywords};
            if(keywords.length === 0) return;
            instance.mark(keywords, {{ className: "keyword-highlight" }});
        }})();
        """
        driver.execute_script(js)
    except Exception as e:
        messagebox.showerror("错误", f"高亮关键词时出错: {str(e)}")

def remove_highlight():
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

def salary_highlight():
    global driver, current_tab
    if driver is None:
        print("❌ 錯誤：請先連接瀏覽器！")
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

def switch_to_tab_by_title(title_contains):
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

def inject_markjs_and_css(driver):
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

def mark_duplicate_words():
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
        
        # 只處理頁面的部分內容（前20000字）以減少處理量
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

        # jieba 分詞（限制處理時間最多5秒）
        start_time = time.time()
        words = list(jieba.cut(text))
        words = [w.strip() for w in words if w.strip()]
        if time.time() - start_time > 5:
            print("jieba分詞超時")
        
        # n-gram 統計（減少範圍到2-4字）
        ngram_counter = Counter()
        start_time = time.time()
        for n in range(2, 5):  # 原本是 2-7，現在縮小到 2-4
            for i in range(len(text) - n + 1):
                if time.time() - start_time > 5:  # 超過5秒就中斷
                    print("n-gram統計超時，已處理部分內容")
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

        # 找出出現超過5次的詞語，且長度大於1，且不在排除清單，且不是薪資組合
        # 縮小上限到500個
        duplicates = [
            w for w, c in word_count.items()
            if c > 5 and len(w) > 1 and w not in exclude_words and not is_number_word(w)
        ]

        # 調試輸出，查看被標記的重複詞
        print("被標記的重複詞:", duplicates)

        # 依長度排序，最多只取前500個
        duplicates = sorted(set(duplicates), key=lambda x: -len(x))[:500]

        if not duplicates:
            messagebox.showinfo("提示", "沒有重複超過5次的詞語")
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

        # 分批標記，每批100個
        batch_size = 100  # 原本是200，現在縮小到100
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
            # 每批標記後稍等一下，避免瀏覽器過載
            time.sleep(0.5)

        messagebox.showinfo("完成", f"已標記 {len(duplicates)} 個重複詞語")
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        messagebox.showerror("錯誤", f"標記重複詞時發生錯誤：\n{str(e)}")

# 添加一个关闭浏览器的函数
def close_browser():
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

# ------------------------------
# GUI 介面
# ------------------------------
root = tk.Tk()
root.title("104 求職網頁工具")

btn_connect_browser = tk.Button(root, text="啟動新的瀏覽器", width=30, command=connect_to_new_browser)
btn_connect_browser.grid(row=0, column=0, padx=5, pady=5)

btn_test_connection = tk.Button(root, text="測試連接", width=30, command=test_connection)
btn_test_connection.grid(row=0, column=1, padx=5, pady=5)

btn_start_scroll = tk.Button(root, text="自動end", width=20, command=start_auto_scroll)
btn_start_scroll.grid(row=1, column=0, padx=5, pady=5)

btn_stop_scroll = tk.Button(root, text="停止end", width=20, command=stop_auto_scroll)
btn_stop_scroll.grid(row=1, column=1, padx=5, pady=5)

btn_highlight = tk.Button(root, text="highlight 關鍵字", width=20, command=highlight_keywords)
btn_highlight.grid(row=2, column=0, padx=5, pady=5)

btn_remove_highlight = tk.Button(root, text="移除highlight", width=20, command=remove_highlight)
btn_remove_highlight.grid(row=2, column=1, padx=5, pady=5)

btn_salary_highlight = tk.Button(root, text="薪資高亮", width=20, command=salary_highlight)
btn_salary_highlight.grid(row=3, column=0, padx=5, pady=5)

btn_mark_duplicate = tk.Button(root, text="標記重複詞", width=20, command=mark_duplicate_words)
btn_mark_duplicate.grid(row=3, column=1, padx=5, pady=5)

btn_close_browser = tk.Button(root, text="關閉瀏覽器", width=20, command=close_browser)
btn_close_browser.grid(row=4, column=0, padx=5, pady=5)

# 窗口关闭时自动清理资源
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
