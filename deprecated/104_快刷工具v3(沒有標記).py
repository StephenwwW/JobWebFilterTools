# 安裝套件 pip install selenium psutil markupsafe

import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import threading
import psutil
import subprocess

# 全局變數
driver = None
current_tab = None
highlighted_keywords = []

def connect_to_new_browser():
    """啟動新的 Chrome 並使用獨立登入資料夾"""
    global driver
    try:
        chrome_options = Options()

        # ✅ 使用獨立的使用者資料資料夾
        chrome_options.add_argument(r'--user-data-dir=C:\Chrome104WebBot')

        chrome_options.add_argument("--disable-infobars")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-extensions")

        service = Service(executable_path="chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        messagebox.showinfo("提示", "已啟動新的 Chrome，請手動登入 104，之後將自動保留登入狀態。")
    except Exception as e:
        messagebox.showerror("錯誤", f"無法啟動 Chrome：{str(e)}")

def start_auto_scroll():
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
    global scrolling
    scrolling = False

def highlight_keywords():
    global highlighted_keywords
    if driver is None:
        messagebox.showerror("錯誤", "請先連接瀏覽器！")
        return
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
        document.querySelectorAll('mark.keyword-highlight').forEach(el => {{
            el.outerHTML = el.innerHTML;
        }});
        var keywords = {highlighted_keywords};
        if(keywords.length === 0) return;
        instance.mark(keywords, {{ className: "keyword-highlight" }});
    }})();
    """
    driver.execute_script(js)

def remove_highlight():
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
    highlighted_keywords = []

def salary_highlight():
    if driver is None:
        print("❌ 錯誤：請先連接瀏覽器！")
        return
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

def highlight_salary_on_page(base_salary, highlight_above, highlight_range, highlight_negotiable):
    inject_markjs_and_css(driver)
    js = f"""
    (function(){{
        if(typeof Mark === 'undefined'){{
            console.error("❌ Mark.js 未載入");
            return;
        }}
        var instance = new Mark(document.body);
        instance.unmark();

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

def switch_to_tab_by_title(title_contains):
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

btn_connect_browser = tk.Button(root, text="啟動新的瀏覽器", width=30, command=connect_to_new_browser)
btn_connect_browser.grid(row=0, column=0, padx=5, pady=5)

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

root.mainloop()
