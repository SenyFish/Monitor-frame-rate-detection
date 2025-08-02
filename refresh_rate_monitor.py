import time
import threading
import json
import os
from datetime import datetime
# from win10toast import ToastNotifier
from plyer import notification
import pystray
from pystray import MenuItem as item
from PIL import Image
import win32api
import win32con
import tkinter as tk
from tkinter import scrolledtext

DATA_FILE = "refresh_rate_stats.json"

def get_refresh_rate():
    try:
        devmode = win32api.EnumDisplaySettings(None, win32con.ENUM_CURRENT_SETTINGS)
        return devmode.DisplayFrequency
    except Exception as e:
        print("获取刷新率失败：", e)
        return -1

def load_stats():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_stats(stats):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

def update_stats(stats, rate=None):
    today = datetime.now().strftime("%Y-%m-%d")
    week = datetime.now().strftime("%Y-W%U")
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    stats.setdefault("daily", {}).setdefault(today, 0)
    stats.setdefault("weekly", {}).setdefault(week, 0)
    stats.setdefault("logs", [])
    stats["daily"][today] += 1
    stats["weekly"][week] += 1
    if rate is not None:
        stats["logs"].append(f"{now_str} 变动为{rate}Hz")
    else:
        stats["logs"].append(now_str)
    save_stats(stats)

def show_stats(stats):
    today = datetime.now().strftime("%Y-%m-%d")
    week = datetime.now().strftime("%Y-W%U")
    daily = stats.get("daily", {}).get(today, 0)
    weekly = stats.get("weekly", {}).get(week, 0)
    return f"今日变动次数: {daily}\n本周变动次数: {weekly}"

def monitor_refresh_rate(stop_event):
    stats = load_stats()
    last_rate = get_refresh_rate()
    while not stop_event.is_set():
        time.sleep(1)
        current_rate = get_refresh_rate()
        if current_rate != last_rate:
            notification.notify(title="刷新率变动", message=f"检测到刷新率变为 {current_rate}Hz", timeout=3)
            update_stats(stats, rate=current_rate)
            last_rate = current_rate

def on_quit(icon, item):
    icon.stop()

def show_stats_gui():
    stats = load_stats()
    today = datetime.now().strftime("%Y-%m-%d")
    week = datetime.now().strftime("%Y-W%U")
    daily = stats.get("daily", {}).get(today, 0)
    weekly = stats.get("weekly", {}).get(week, 0)
    root = tk.Tk()
    root.title("刷新率变动统计")
    root.geometry("300x150")
    label = tk.Label(root, text=f"今日变动次数: {daily}\n本周变动次数: {weekly}", font=("微软雅黑", 14))
    label.pack(expand=True, fill='both', padx=10, pady=10)
    root.attributes('-topmost', True)
    root.mainloop()

def on_show_stats(icon, item):
    import threading
    threading.Thread(target=show_stats_gui, daemon=True).start()

def show_logs_gui():
    stats = load_stats()
    logs = stats.get("logs", [])
    root = tk.Tk()
    root.title("刷新率变动记录")
    root.geometry("400x300")
    st = scrolledtext.ScrolledText(root, wrap=tk.WORD)
    st.pack(expand=True, fill='both')
    if not logs:
        st.insert(tk.END, "暂无变动记录")
    else:
        for log in logs:
            st.insert(tk.END, log + "\n")
    st.config(state=tk.DISABLED)
    root.attributes('-topmost', True)
    root.mainloop()

def on_show_logs(icon, item):
    import threading
    threading.Thread(target=show_logs_gui, daemon=True).start()

def main():
    stop_event = threading.Event()
    t = threading.Thread(target=monitor_refresh_rate, args=(stop_event,), daemon=True)
    t.start()

    # 创建蓝色和红色图标
    blue_image = Image.new('RGB', (64, 64), color=(0, 128, 255))
    red_image = Image.new('RGB', (64, 64), color=(255, 0, 0))

    # 初始刷新率决定初始图标
    try:
        initial_rate = get_refresh_rate()
    except:
        initial_rate = 0
    icon_image = blue_image if initial_rate >= 180 else red_image

    icon = pystray.Icon("refresh_rate_monitor", icon_image, "刷新率监控", menu=pystray.Menu(
        item('显示统计', on_show_stats),
        item('查看变动记录', on_show_logs),
        item('退出', on_quit)
    ))

    def update_icon():
        while not stop_event.is_set():
            rate = get_refresh_rate()
            if rate >= 180:
                icon.icon = blue_image
            else:
                icon.icon = red_image
            time.sleep(1)

    threading.Thread(target=update_icon, daemon=True).start()
    icon.run()
    stop_event.set()

if __name__ == "__main__":
    main()