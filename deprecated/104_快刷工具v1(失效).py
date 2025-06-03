# 安裝套件 pip install selenium psutil markupsafe

# 安裝套件 pip install -U selenium


import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import threading
import subprocess
import psutil
import time
from selenium import webdriver

def start_chrome_debugging():
    """檢查 Chrome 是否已開啟 remote-debugging-port，若無則自動啟動"""
    chrome_running = any("chrome.exe" in p.name().lower() for p in psutil.process_iter(attrs=["pid", "name"]))
    
    if not chrome_running:
        subprocess.Popen(r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222', shell=True)
        time.sleep(3)  # 等待 Chrome 啟動

def connect_to_existing_browser():
    """連接到已經開啟的瀏覽器"""
    global driver
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        driver = webdriver.Chrome(options=options)
        print("✅ 成功連接到已開啟的瀏覽器！")
    except Exception as e:
        print(f"❌ 錯誤：無法連接到瀏覽器：{str(e)}")

# 自動啟動 Chrome 並連接
start_chrome_debugging()
connect_to_existing_browser()


# 全局變數，存儲 driver 和當前的分頁
driver = None
current_tab = None

# 存儲高亮的關鍵字
highlighted_keywords = []

# ------------------------------
# 連接已開啟的瀏覽器
# ------------------------------
def connect_to_existing_browser():
    """連接到已經開啟的瀏覽器窗口"""
    global driver
    try:
        # 使用遠程調試連接
        options = webdriver.ChromeOptions()
        options.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
        driver = webdriver.Chrome(options=options)
        messagebox.showinfo("提示", "成功連接到已開啟的瀏覽器！")
        switch_to_tab_by_title("104")  # 這裡假設你開啟的是104的分頁，根據需要改動
    except Exception as e:
        messagebox.showerror("錯誤", f"無法連接到瀏覽器：{str(e)}")

# ------------------------------
# 控制功能
# ------------------------------
def start_auto_scroll():
    """開始自動捲動，持續執行捲動至頁尾"""
    global scrolling
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    scrolling = True

    def scroll():
        while scrolling:
            try:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            except Exception as e:
                print("捲動錯誤:", e)
            time.sleep(1)

    t = threading.Thread(target=scroll)
    t.daemon = True
    t.start()

def stop_auto_scroll():
    """停止自動捲動"""
    global scrolling
    scrolling = False

def highlight_keywords():
    """使用 mark.js 高亮使用者指定的關鍵字"""
    global highlighted_keywords
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    inject_markjs_and_css(driver)
    keywords = simpledialog.askstring("關鍵字", "請輸入關鍵字 (以逗號分隔):")
    if not keywords:
        return
    # 保留舊的高亮，並添加新高亮
    highlighted_keywords += [keyword.strip() for keyword in keywords.split(",")]

    js = f"""
    (function(){{
        var instance = new Mark(document.body);
        instance.unmark();  // 先清除之前的標記
        var keywords = {highlighted_keywords};
        if(keywords.length === 0) return;
        instance.mark(keywords);
    }})();
    """
    driver.execute_script(js)

def remove_highlight():
    """移除所有 mark.js 所做的高亮"""
    global highlighted_keywords
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    inject_markjs_and_css(driver)
    js = """
    (function(){
        var instance = new Mark(document.body);
        instance.unmark();
    })();
    """
    driver.execute_script(js)
    highlighted_keywords = []  # 清空已高亮的關鍵字

def salary_highlight():
    """設定薪資篩選條件，並高亮符合的薪資範圍"""
    if driver is None:
        print("❌ 錯誤：請先連接瀏覽器！")
        return

    # 創建 UI 視窗
    dialog = tk.Toplevel()
    dialog.title("薪資篩選設定")

    tk.Label(dialog, text="請輸入基準薪資：").grid(row=0, column=0, padx=10, pady=5)
    salary_var = tk.IntVar(value=30000)
    entry_salary = tk.Entry(dialog, textvariable=salary_var, width=10)
    entry_salary.grid(row=0, column=1, padx=10, pady=5)

    # 勾選框
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

def highlight_salary_on_page(base_salary, highlight_above, highlight_range, highlight_negotiable):
    """使用 JavaScript 高亮薪資範圍"""
    inject_markjs_and_css(driver)  # **先確保 Mark.js 已載入**

    js = f"""
    (function(){{
        if(typeof Mark === 'undefined'){{
            console.error("❌ Mark.js 未載入，無法執行高亮！");
            return;
        }}

        var instance = new Mark(document.body);
        instance.unmark();  // 清除先前的標記

        var baseSalary = {base_salary};
        var salaryPattern = /月薪\\s*([\\d,]+)(?:~([\\d,]+))?\\s*元/g;
        var bodyText = document.body.innerHTML;

        var matches = [...bodyText.matchAll(salaryPattern)];
        var highlightTargets = [];

        matches.forEach(match => {{
            var minSalary = parseInt(match[1].replace(/,/g, ''), 10);
            var maxSalary = match[2] ? parseInt(match[2].replace(/,/g, ''), 10) : null;

            // 檢查基準薪資以上
            if ({'true' if highlight_above else 'false'} && minSalary >= baseSalary) {{
                highlightTargets.push(match[1]);
            }}

            // 檢查基準薪資區間
            if ({'true' if highlight_range else 'false'} && minSalary <= baseSalary && maxSalary >= baseSalary) {{
                highlightTargets.push(match[2]);
            }}
        }});

        // 額外高亮「待遇面議」
        if ({'true' if highlight_negotiable else 'false'}) {{
            highlightTargets.push("待遇面議");
        }}

        if (highlightTargets.length > 0) {{
            instance.mark(highlightTargets.filter(Boolean));
        }}
    }})();
    """
    driver.execute_script(js)

def switch_to_tab_by_title(title_contains):
    """切換到指定標題包含某個字串的分頁"""
    global driver, current_tab
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
    handles = driver.window_handles
    for handle in handles:
        driver.switch_to.window(handle)
        if title_contains in driver.title:
            current_tab = handle
            messagebox.showinfo("提示", f"已切換到標題包含 '{title_contains}' 的分頁")
            return
    messagebox.showwarning("警告", f"未找到標題包含 '{title_contains}' 的分頁")

def inject_markjs_and_css(driver):
    """確保 mark.js 與 CSS 已經注入頁面"""
    mark_exists = driver.execute_script("return typeof Mark !== 'undefined';")
    if not mark_exists:
        # 注入 mark.js
        js_inject = """
        (function(){
            if(typeof Mark === 'undefined'){
                var script = document.createElement('script');
                script.src = "https://cdnjs.cloudflare.com/ajax/libs/mark.js/8.11.1/mark.min.js";
                script.onload = function(){
                    console.log("✅ mark.js 已載入");
                };
                document.head.appendChild(script);
            }
        })();
        """
        driver.execute_script(js_inject)

        # **等待 mark.js 載入完成**
        for _ in range(10):  # 最多等 10 秒
            mark_exists = driver.execute_script("return typeof Mark !== 'undefined';")
            if mark_exists:
                break
            time.sleep(1)

    # 注入自訂 CSS
    css_inject = """
    (function(){
        if(!document.getElementById("custom-highlight-style")){
            var style = document.createElement('style');
            style.id = "custom-highlight-style";
            style.innerHTML = `
                mark { background: yellow; color: black; }
                mark.salary-highlight { background: orange; color: black; }
            `;
            document.head.appendChild(style);
        }
    })();
    """
    driver.execute_script(css_inject)


# ------------------------------
# GUI 介面
# ------------------------------
root = tk.Tk()
root.title("104 求職網頁工具")

# 第一列：連接到已開啟的瀏覽器
btn_connect_browser = tk.Button(root, text="連接已開啟的瀏覽器", width=30, command=connect_to_existing_browser)
btn_connect_browser.grid(row=0, column=0, padx=5, pady=5)

# 第二列：自動捲動控制
btn_start_scroll = tk.Button(root, text="自動end", width=20, command=start_auto_scroll)
btn_start_scroll.grid(row=1, column=0, padx=5, pady=5)

btn_stop_scroll = tk.Button(root, text="停止end", width=20, command=stop_auto_scroll)
btn_stop_scroll.grid(row=1, column=1, padx=5, pady=5)

# 第三列：關鍵字高亮
btn_highlight = tk.Button(root, text="highlight 關鍵字", width=20, command=highlight_keywords)
btn_highlight.grid(row=2, column=0, padx=5, pady=5)

btn_remove_highlight = tk.Button(root, text="移除highlight", width=20, command=remove_highlight)
btn_remove_highlight.grid(row=2, column=1, padx=5, pady=5)

# 薪資高亮選擇功能
btn_salary_highlight = tk.Button(root, text="薪資高亮", width=20, command=salary_highlight)
btn_salary_highlight.grid(row=3, column=0, padx=5, pady=5)

root.mainloop()
