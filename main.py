# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
import time
import random
from datetime import datetime, timedelta
import threading
import configparser
import os
import queue
import ctypes
import ctypes.wintypes as wt
import traceback

APP_TITLE = "АвтоДоповідь WhatsApp — стабільна"
SELF_PID = os.getpid()

CONFIG_FILE = "report.ini"
CONFIG_SECTION = "Report"
CONFIG_KEY = "text"

ANTIFLOOD_SECONDS = 15 * 60  # 15 хв
VERIFY_BEFORE_SEND = True
VERIFY_RETRIES = 3

PYAUTOGUI_AVAILABLE = False
PYPERCLIP_AVAILABLE = False
PYWINAUTO_AVAILABLE = False
PSUTIL_AVAILABLE = False

try:
    import pyautogui
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.05
    PYAUTOGUI_AVAILABLE = True
except Exception as e:
    print(f"⚠️ pyautogui недоступний: {e}")

try:
    import pyperclip
    PYPERCLIP_AVAILABLE = True
except Exception as e:
    print(f"⚠️ pyperclip недоступний: {e}")

try:
    from pywinauto.application import Application
    from pywinauto import findwindows
    from pywinauto.keyboard import send_keys
    PYWINAUTO_AVAILABLE = True
except Exception as e:
    print(f"ℹ️ pywinauto не встановлено (рекомендовано): {e}")

try:
    import psutil
    PSUTIL_AVAILABLE = True
except Exception:
    PSUTIL_AVAILABLE = False

# ---------------------- WinAPI helpers ----------------------
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wt.HWND, wt.LPARAM)
GetWindowTextW = user32.GetWindowTextW
GetWindowTextW.argtypes = [wt.HWND, wt.LPWSTR, ctypes.c_int]
GetWindowTextW.restype = ctypes.c_int

IsWindowVisible = user32.IsWindowVisible
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowRect = user32.GetWindowRect
SetForegroundWindow = user32.SetForegroundWindow
ShowWindow = user32.ShowWindow

SW_RESTORE = 9

def get_window_title(hwnd):
    buf = ctypes.create_unicode_buffer(512)
    GetWindowTextW(hwnd, buf, 512)
    return buf.value.strip()

def get_window_pid(hwnd):
    pid = wt.DWORD()
    GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value

def is_window_visible(hwnd):
    return bool(IsWindowVisible(hwnd))

def enum_visible_top_windows():
    result = []
    def callback(hwnd, lParam):
        try:
            if not is_window_visible(hwnd):
                return True
            title = get_window_title(hwnd)
            if not title:
                return True
            result.append(hwnd)
        except Exception:
            pass
        return True
    user32.EnumWindows(EnumWindowsProc(callback), 0)
    return result

def get_window_rect(hwnd):
    rect = wt.RECT()
    GetWindowRect(hwnd, ctypes.byref(rect))
    return rect

def restore_and_foreground(hwnd):
    try:
        ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.1)
        SetForegroundWindow(hwnd)
        return True
    except Exception:
        return False

# ---------------------- Глобальний стан ----------------------
state_lock = threading.Lock()
timer_active = False
next_report_time = None
last_fired_target = None
last_send_ts = 0.0
timer_thread = None
fire_lock = threading.Lock()

log_q = queue.Queue()
def log_message(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    log_q.put(f"[{ts}] {msg}\n")

def log_exception(prefix: str, e: Exception):
    tb = traceback.format_exc(limit=2)
    log_message(f"{prefix}: {e.__class__.__name__}: {e}")
    log_message(f"↳ Trace: {tb.strip()}")

# ---------------------- Утиліти часу ----------------------
def load_saved_text():
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        cfg.read(CONFIG_FILE, encoding="utf-8-sig")
        return cfg.get(CONFIG_SECTION, CONFIG_KEY, fallback="")
    return ""

def save_text(text):
    cfg = configparser.ConfigParser()
    cfg[CONFIG_SECTION] = {CONFIG_KEY: text}
    with open(CONFIG_FILE, "w", encoding="utf-8-sig") as f:
        cfg.write(f)

def get_next_slot(base=None):
    now = base or datetime.now()
    if now.minute < 45:
        t = now.replace(minute=45, second=0, microsecond=0)
    else:
        t = (now + timedelta(hours=1)).replace(minute=45, second=0, microsecond=0)
    offset = random.randint(-2, 2)
    return t + timedelta(minutes=offset)

def get_next_hour_slot_from_target(prev_target):
    base = (prev_target + timedelta(hours=1)).replace(minute=45, second=0, microsecond=0)
    offset = random.randint(-2, 2)
    return base + timedelta(minutes=offset)

# ---------------------- Пошук вікна WhatsApp ----------------------
BROWSER_NAMES = ("chrome", "msedge", "firefox", "opera", "opera_gx", "vivaldi", "brave")

def list_candidate_pids():
    """PIDs для WhatsApp Desktop + браузери."""
    pids_whatsapp = set()
    pids_browsers = set()
    if not PSUTIL_AVAILABLE:
        return pids_whatsapp, pids_browsers

    for p in psutil.process_iter(['pid', 'name']):
        name = (p.info.get('name') or "").lower()
        if "whatsapp" in name:
            pids_whatsapp.add(p.info['pid'])
        elif any(b in name for b in BROWSER_NAMES):
            pids_browsers.add(p.info['pid'])
    return pids_whatsapp, pids_browsers

def find_whatsapp_window():
    """
    Повертає кортеж (hwnd, origin) де origin ∈ {"desktop","web","unknown"} або (None, None).
    Спочатку шукаємо вікна процесів WhatsApp, далі — браузери з title, що містить "WhatsApp".
    Відсікаємо наше Tk-вікно за SELF_PID та APP_TITLE.
    """
    hwnds = enum_visible_top_windows()
    pids_whatsapp, pids_browsers = list_candidate_pids()
    desktop_best = None
    web_best = None

    for hwnd in hwnds:
        try:
            pid = get_window_pid(hwnd)
            title = get_window_title(hwnd)
            if not title:
                continue
            # відсікаємо наше вікно
            if pid == SELF_PID or (APP_TITLE and APP_TITLE in title):
                continue

            # 1) WhatsApp Desktop по PID процесу
            if pid in pids_whatsapp:
                desktop_best = hwnd
                # якщо в заголовку є "WhatsApp" — це майже те, що треба
                if "whatsapp" in title.lower():
                    return hwnd, "desktop"
                continue

            # 2) WhatsApp Web у браузері: браузерний PID + title містить 'WhatsApp'
            if pid in pids_browsers and ("whatsapp" in title.lower() or "web.whatsapp" in title.lower()):
                if web_best is None:
                    web_best = hwnd
        except Exception:
            continue

    if desktop_best:
        return desktop_best, "desktop"
    if web_best:
        return web_best, "web"
    # як останній шанс: будь-яке видиме вікно з 'WhatsApp' у заголовку
    for hwnd in hwnds:
        title = get_window_title(hwnd)
        pid = get_window_pid(hwnd)
        if pid == SELF_PID:
            continue
        if "whatsapp" in title.lower():
            return hwnd, "unknown"
    return None, None

# ---------------------- Вставка у WhatsApp ----------------------
def _uia_set_focus_and_type(text: str, do_send: bool) -> bool:
    if not PYWINAUTO_AVAILABLE:
        return False
    hwnd, origin = find_whatsapp_window()
    if not hwnd:
        log_message("UIA: не знайдено вікна WhatsApp.")
        return False
    try:
        restore_and_foreground(hwnd)
        app = Application(backend="uia").connect(handle=hwnd, timeout=5)
        win = app.window(handle=hwnd)
        try:
            edits = win.descendants(control_type="Edit")
        except Exception:
            edits = []
        if not edits:
            log_message("UIA: не знайдено поле вводу (Edit).")
            return False

        # беремо найнижче поле
        try:
            edits.sort(key=lambda ed: ed.rectangle().bottom, reverse=True)
        except Exception:
            pass
        edit = edits[0]
        try: edit.set_focus()
        except Exception: pass
        try: edit.click_input()
        except Exception: pass
        time.sleep(0.08)

        safe_text = text.replace(" ", "{SPACE}")
        log_message("UIA: друкую текст (send_keys)…")
        send_keys(safe_text, with_newlines=True, pause=0.01)

        if VERIFY_BEFORE_SEND:
            if not _verify_via_clipboard(text):
                log_message("UIA: верифікація не пройшла (текст у полі не збігся).")
                return False
        if do_send:
            send_keys("{ENTER}")
        return True
    except Exception as e:
        log_exception("UIA: помилка друку", e)
        return False

def _uia_focus_and_paste(text: str, do_send: bool, pre_ms: int, paste_delay_s: float) -> bool:
    if not (PYWINAUTO_AVAILABLE and PYAUTOGUI_AVAILABLE and PYPERCLIP_AVAILABLE):
        return False
    hwnd, origin = find_whatsapp_window()
    if not hwnd:
        log_message("UIA+Paste: не знайдено вікна WhatsApp.")
        return False
    try:
        restore_and_foreground(hwnd)
        app = Application(backend="uia").connect(handle=hwnd, timeout=5)
        win = app.window(handle=hwnd)
        try:
            edits = win.descendants(control_type="Edit")
        except Exception:
            edits = []
        if not edits:
            log_message("UIA+Paste: не знайдено поле вводу (Edit).")
            return False

        try:
            edits.sort(key=lambda ed: ed.rectangle().bottom, reverse=True)
        except Exception:
            pass
        edit = edits[0]
        try: edit.set_focus()
        except Exception: pass
        try: edit.click_input()
        except Exception: pass

        time.sleep(max(0.05, pre_ms/1000.0))
        pyperclip.copy(text)
        time.sleep(0.12)
        log_message("UIA+Paste: Ctrl+V…")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(max(0.05, paste_delay_s))

        if VERIFY_BEFORE_SEND:
            if not _verify_via_clipboard(text):
                log_message("UIA+Paste: верифікація не пройшла (текст у полі не збігся).")
                return False
        if do_send:
            send_keys("{ENTER}")
        return True
    except Exception as e:
        log_exception("UIA+Paste: помилка вставки", e)
        return False

def _pgui_click_and_paste(text: str, do_send: bool, pre_ms: int, paste_delay_s: float, send_delay_s: float) -> bool:
    if not (PYAUTOGUI_AVAILABLE and PYPERCLIP_AVAILABLE):
        return False
    hwnd, origin = find_whatsapp_window()
    if not hwnd:
        log_message("PyAutoGUI: не знайдено вікна WhatsApp.")
        return False
    try:
        restore_and_foreground(hwnd)
        rect = get_window_rect(hwnd)
        # Клік у нижню частину (поле вводу)
        cx = rect.left + (rect.right - rect.left)//2
        cy = rect.bottom - 60
        pyautogui.press('esc')
        time.sleep(0.05)
        pyautogui.click(cx, cy)
        time.sleep(max(0.0, pre_ms/1000.0))

        pyperclip.copy(text)
        time.sleep(0.12)
        log_message("PyAutoGUI: Ctrl+V…")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(max(0.05, paste_delay_s))

        if VERIFY_BEFORE_SEND:
            if not _verify_via_clipboard(text):
                log_message("PyAutoGUI: верифікація не пройшла (текст у полі не збігся).")
                pyautogui.press('end')
                return False
            pyautogui.press('end')

        if do_send:
            pyautogui.press('enter')
            time.sleep(max(0.05, send_delay_s))
        return True
    except Exception as e:
        log_exception("PyAutoGUI: помилка вставки", e)
        return False

def _verify_via_clipboard(expected: str) -> bool:
    """Ctrl+A → Ctrl+C → порівняння з expected; повертаємо курсор у кінець."""
    try:
        if PYAUTOGUI_AVAILABLE:
            pyautogui.hotkey('ctrl', 'a'); time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'c'); time.sleep(0.08)
            if PYPERCLIP_AVAILABLE:
                got = pyperclip.paste()
                ok = (got == expected)
            else:
                ok = False
            pyautogui.press('end')
            return ok
        else:
            # UIA клавіатурою
            send_keys("^a"); time.sleep(0.05)
            send_keys("^c"); time.sleep(0.08)
            got = pyperclip.paste() if PYPERCLIP_AVAILABLE else None
            send_keys("{END}")
            return got == expected
    except Exception as e:
        log_exception("Верифікація", e)
        return False

def whatsapp_send(text: str, do_send=True, pre_ms=300, paste_delay_s=0.5, send_delay_s=0.3) -> bool:
    methods = [
        ("UIA: друк", lambda: _uia_set_focus_and_type(text, do_send)),
        ("UIA: Ctrl+V", lambda: _uia_focus_and_paste(text, do_send, pre_ms, paste_delay_s)),
        ("PyAutoGUI: Ctrl+V", lambda: _pgui_click_and_paste(text, do_send, pre_ms, paste_delay_s, send_delay_s)),
    ]
    for name, fn in methods:
        for attempt in range(1, VERIFY_RETRIES + 1):
            log_message(f"→ Спроба [{name}] #{attempt}…")
            ok = fn()
            if ok:
                log_message(f"✅ Успіх методом [{name}]")
                return True
            else:
                log_message(f"⚠️ [{name}] не вдалась (спроба {attempt}).")
                time.sleep(0.25 * attempt)
    return False

# ---------------------- Відправка/таймер ----------------------
def do_send_report(text, pre_ms, paste_s, send_s, via_timer=False):
    global last_send_ts
    now_ts = time.time()
    with state_lock:
        if now_ts - last_send_ts < ANTIFLOOD_SECONDS:
            left = int(ANTIFLOOD_SECONDS - (now_ts - last_send_ts))
            log_message(f"⛔ Скасовано дубль: антифлуд {ANTIFLOOD_SECONDS//60} хв. Залишилось ~{left}с.")
            return
        last_send_ts = now_ts

    prefix = "⏰ [Таймер] " if via_timer else ""
    text = (text or "").strip()
    if not text:
        log_message(prefix + "⚠️ Текст порожній.")
        return
    log_message(prefix + "📤 Відправляю…")
    ok = whatsapp_send(text, True, pre_ms, paste_s, send_s)
    if ok:
        log_message(prefix + "🎉 Готово.")
    else:
        log_message(prefix + "❌ Не вдалося вставити/відправити.")

def schedule_thread():
    global next_report_time, timer_active, last_fired_target
    with state_lock:
        timer_active = True
        if next_report_time is None:
            next_report_time = get_next_slot()
    log_message("✅ Таймер запущено.")

    while True:
        with state_lock:
            active = timer_active
            target = next_report_time
            fired_for_target = (last_fired_target == target)
        if not active:
            break

        now = datetime.now()
        if now >= target and not fired_for_target:
            if not fire_lock.acquire(blocking=False):
                time.sleep(0.1)
                continue
            try:
                log_message(f"⏰ ТАЙМЕР: {target.strftime('%H:%M:%S')} — відправляю автоматично.")
                with state_lock:
                    last_fired_target = target
                def read_and_dispatch():
                    t = entry.get()
                    pre = pre_paste_delay.get()
                    pd = paste_delay.get()
                    sd = send_delay.get()
                    threading.Thread(
                        target=do_send_report,
                        args=(t, pre, pd, sd, True),
                        daemon=True
                    ).start()
                root.after(0, read_and_dispatch)
                with state_lock:
                    next_report_time = get_next_hour_slot_from_target(target)
                    log_message(f"📅 Наступна доповідь запланована на {next_report_time.strftime('%H:%M:%S')}")
            finally:
                fire_lock.release()
        time.sleep(0.2)

# ---------------------- GUI ----------------------
root = tk.Tk()
root.title(APP_TITLE)
root.geometry("940x760")

paste_delay = tk.DoubleVar(value=0.8)
send_delay = tk.DoubleVar(value=0.3)
pre_paste_delay = tk.IntVar(value=300)

notebook = ttk.Notebook(root); notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

main_frame = ttk.Frame(notebook); notebook.add(main_frame, text="Основні")
tk.Label(main_frame, text="Введіть текст доповіді:", font=("Arial", 12, "bold")).pack(pady=10)
entry = tk.Entry(main_frame, width=70, font=("Arial", 11))
entry.insert(0, load_saved_text()); entry.pack(pady=5)
entry.bind("<KeyRelease>", lambda e: save_text(entry.get()))

delay_frame = tk.LabelFrame(main_frame, text="Налаштування затримок", font=("Arial", 10, "bold"))
delay_frame.pack(pady=15, padx=20, fill=tk.X)

r1 = tk.Frame(delay_frame); r1.pack(fill=tk.X, padx=10, pady=5)
tk.Label(r1, text="Затримка перед вставкою:", font=("Arial", 10)).pack(side=tk.LEFT)
tk.Spinbox(r1, from_=0, to=5000, increment=50, textvariable=pre_paste_delay, width=10, font=("Arial", 10)).pack(side=tk.RIGHT)
tk.Label(r1, text="мс", font=("Arial", 10)).pack(side=tk.RIGHT, padx=(0,5))

r2 = tk.Frame(delay_frame); r2.pack(fill=tk.X, padx=10, pady=5)
tk.Label(r2, text="Затримка після вставки:", font=("Arial", 10)).pack(side=tk.LEFT)
tk.Spinbox(r2, from_=0.1, to=5.0, increment=0.1, textvariable=paste_delay, width=10, font=("Arial", 10)).pack(side=tk.RIGHT)
tk.Label(r2, text="секунд", font=("Arial", 10)).pack(side=tk.RIGHT, padx=(0,5))

r3 = tk.Frame(delay_frame); r3.pack(fill=tk.X, padx=10, pady=5)
tk.Label(r3, text="Затримка після відправки:", font=("Arial", 10)).pack(side=tk.LEFT)
tk.Spinbox(r3, from_=0.1, to=5.0, increment=0.1, textvariable=send_delay, width=10, font=("Arial", 10)).pack(side=tk.RIGHT)
tk.Label(r3, text="секунд", font=("Arial", 10)).pack(side=tk.RIGHT, padx=(0,5))

timer_frame = tk.LabelFrame(main_frame, text="Автоматичні доповіді", font=("Arial", 10, "bold"))
timer_frame.pack(pady=15, padx=20, fill=tk.X)

btns = tk.Frame(timer_frame); btns.pack(pady=10)
def start_timer():
    global timer_active, timer_thread
    with state_lock:
        if timer_active and timer_thread and timer_thread.is_alive():
            log_message("⚠️ Таймер уже працює (активний тред).")
            return
        timer_active = True
        if next_report_time is None:
            globals()['next_report_time'] = get_next_slot()
    timer_thread = threading.Thread(target=schedule_thread, daemon=True)
    timer_thread.start()
    log_message("▶️ Запуск таймера…")

def stop_timer():
    global timer_active
    with state_lock:
        timer_active = False
    log_message("🛑 Таймер зупинено.")

tk.Button(btns, text="Запустити таймер", command=start_timer, font=("Arial", 10), bg="#4CAF50", fg="white", width=15).pack(side=tk.LEFT, padx=5)
tk.Button(btns, text="Зупинити таймер", command=stop_timer, font=("Arial", 10), bg="#f44336", fg="white", width=15).pack(side=tk.LEFT, padx=5)

timer_label = tk.Label(timer_frame, text="", font=("Arial", 12), fg="#333")
timer_label.pack(pady=10)

actions = tk.LabelFrame(main_frame, text="Дії", font=("Arial", 10, "bold"))
actions.pack(pady=15, padx=20, fill=tk.X)

def send_now():
    t = entry.get()
    pre = pre_paste_delay.get()
    pd = paste_delay.get()
    sd = send_delay.get()
    threading.Thread(target=do_send_report, args=(t, pre, pd, sd, False), daemon=True).start()

def test_insert():
    t = entry.get().strip()
    if not t:
        log_message("⚠️ Текст для тесту порожній.")
        return
    def worker():
        log_message("🧪 Тест: друк/вставка без Enter…")
        ok = whatsapp_send(t, do_send=False, pre_ms=pre_paste_delay.get(),
                           paste_delay_s=paste_delay.get(), send_delay_s=send_delay.get())
        if ok: log_message("🎉 Вставка пройшла (без відправки).")
        else:  log_message("❌ Не вдалося вставити у тесті.")
    threading.Thread(target=worker, daemon=True).start()

def diagnose():
    log_message("🔬 Діагностика:")
    log_message(f"  pywinauto: {PYWINAUTO_AVAILABLE}")
    log_message(f"  pyautogui: {PYAUTOGUI_AVAILABLE}")
    log_message(f"  pyperclip: {PYPERCLIP_AVAILABLE}")
    log_message(f"  psutil: {PSUTIL_AVAILABLE}")

    # Виведемо кандидатні вікна
    hwnd, origin = find_whatsapp_window()
    if hwnd:
        title = get_window_title(hwnd)
        pid = get_window_pid(hwnd)
        rect = get_window_rect(hwnd)
        log_message(f"  Знайдено WhatsApp ({origin}) HWND={hwnd} PID={pid} TITLE='{title}' "
                    f"RECT=({rect.left},{rect.top},{rect.right},{rect.bottom})")
    else:
        log_message("  WhatsApp-вікно не знайдено. Відкрий WhatsApp Desktop або вкладку web.whatsapp у браузері.")

row = tk.Frame(actions); row.pack(pady=10)
tk.Button(row, text="Відправити зараз", command=send_now, font=("Arial", 9), bg="#2196F3", fg="white", width=17).pack(side=tk.LEFT, padx=3)
tk.Button(row, text="Тест вставлення", command=test_insert, font=("Arial", 9), bg="#FF9800", fg="white", width=17).pack(side=tk.LEFT, padx=3)
tk.Button(row, text="Діагностика", command=diagnose, font=("Arial", 9), bg="#9C27B0", fg="white", width=17).pack(side=tk.LEFT, padx=3)

log_tab = ttk.Frame(notebook); notebook.add(log_tab, text="Логи")
log_header = tk.Frame(log_tab); log_header.pack(fill=tk.X, padx=10, pady=5)
tk.Label(log_header, text="Логи:", font=("Arial", 12, "bold")).pack(side=tk.LEFT)
def clear_log():
    log_text.delete(1.0, tk.END)
tk.Button(log_header, text="Очистити", command=clear_log, font=("Arial", 10), bg="#607D8B", fg="white").pack(side=tk.RIGHT)

log_text = tk.Text(log_tab, wrap=tk.WORD, font=("Consolas", 10), bg="#f5f5f5", fg="#333")
log_scroll = tk.Scrollbar(log_tab, command=log_text.yview)
log_text.config(yscrollcommand=log_scroll.set)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10,0), pady=10)
log_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0,10), pady=10)

def pump_logs():
    try:
        while True:
            line = log_q.get_nowait()
            log_text.insert(tk.END, line)
            log_text.see(tk.END)
    except queue.Empty:
        pass
    root.after(50, pump_logs)

def compute_display_target():
    with state_lock:
        target = next_report_time
    if target is not None:
        return target
    return get_next_slot()

def update_timer_label():
    target = compute_display_target()
    now = datetime.now()
    remaining = target - now
    if remaining.total_seconds() < 0:
        target = get_next_slot(now + timedelta(seconds=1))
        remaining = target - now
    mins, secs = divmod(int(remaining.total_seconds()), 60)
    hours, mins = divmod(mins, 60)
    with state_lock:
        active = timer_active
    status = "🟢 Таймер активний" if active else "⚪ Таймер вимкнений"
    timer_label.config(
        text=f"{status}\nНаступна доповідь: {target.strftime('%H:%M:%S')}\nЗалишилось: {hours:02d}:{mins:02d}:{secs:02d}"
    )
    root.after(200, update_timer_label)

root.title(APP_TITLE)
log_message("🚀 Запуск. Рекомендовано: pip install pywinauto psutil")

root.after(0, pump_logs)
root.after(0, update_timer_label)
root.mainloop()
