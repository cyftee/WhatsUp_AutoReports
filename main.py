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

# ================== НАЛАШТУВАННЯ/КОНСТАНТИ ==================
APP_TITLE = "АвтоДоповідь WhatsApp — стабільна"
SELF_PID = os.getpid()

CONFIG_FILE = "report.ini"
CONFIG_SECTION = "Report"
CONFIG_KEY = "text"

# Антифлуд: 15 хв (у секундах)
ANTIFLOOD_SECONDS = 15 * 60  # 900s

# Опційні бібліотеки
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

# ================== ГЛОБАЛЬНИЙ СТАН ==================
state_lock = threading.Lock()
timer_active = False
next_report_time = None           # план наступного запуску (лише таймер-тред змінює)
last_fired_target = None          # для якого target уже стріляли
last_send_ts = 0.0                # антифлуд: штамп останньої відправки (секунди)

# єдиний таймер-тред + замок на миттєве спрацювання
timer_thread = None
fire_lock = threading.Lock()

log_q = queue.Queue()
def log_message(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    log_q.put(f"[{ts}] {msg}\n")

# ================== УТИЛІТИ ==================
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
    """
    Найближча 45-та хвилина поточної/наступної години з випадковим офсетом -2..+2 хв.
    Використовуємо для первинного планування та для планування від довільної бази.
    """
    now = base or datetime.now()
    if now.minute < 45:
        t = now.replace(minute=45, second=0, microsecond=0)
    else:
        t = (now + timedelta(hours=1)).replace(minute=45, second=0, microsecond=0)
    offset = random.randint(-2, 2)
    return t + timedelta(minutes=offset)

def get_next_hour_slot_from_target(prev_target):
    """
    Планує наступний запуск ВІД попередньої цілі + 1 година, а не від 'тепер'.
    Це усуває дублювання в межах тієї ж години.
    """
    base = (prev_target + timedelta(hours=1))
    return get_next_slot(base)

# ================== WHATSAPP: ПОШУК/ФОКУС/ВСТАВКА ==================
def _pywinauto_find_main():
    """Знаходимо головне вікно WhatsApp через UIA і відсікаємо наше Tk-вікно."""
    if not PYWINAUTO_AVAILABLE:
        return None
    try:
        handles = findwindows.find_windows(title_re=r".*WhatsApp.*", visible_only=True)
        for h in handles:
            app = Application(backend="uia").connect(handle=h, timeout=5)
            dlg = app.window(handle=h)
            if not (dlg.exists() and dlg.is_visible()):
                continue

            # Відсікаємо наше вікно за PID/класом/титулом
            try:
                pid = dlg.element_info.process_id
            except Exception:
                pid = None
            try:
                cls = dlg.element_info.class_name or ""
            except Exception:
                cls = ""
            try:
                title = dlg.window_text() or ""
            except Exception:
                title = ""

            if pid == SELF_PID:
                continue
            if "TkTopLevel" in cls:
                continue
            if APP_TITLE and APP_TITLE in title:
                continue

            # Опціонально — перевіряємо ім'я процесу
            if PSUTIL_AVAILABLE and pid:
                try:
                    name = psutil.Process(pid).name()
                    if "whatsapp" not in name.lower():
                        continue
                except Exception:
                    pass

            return app, dlg
    except Exception as e:
        log_message(f"pywinauto пошук: {e}")
    return None

def _pywinauto_focus_and_type(text: str, do_send: bool) -> bool:
    """Фокус вікна, пошук поля вводу (Edit) і друк напряму. Пробіли -> {SPACE}."""
    found = _pywinauto_find_main()
    if not found:
        return False
    app, win = found
    try:
        try:
            win.restore()
        except Exception:
            pass
        win.set_focus()
        time.sleep(0.15)

        # Знайти всі Edit (найнижчий — композер)
        try:
            edits = win.descendants(control_type="Edit")
        except Exception:
            edits = []

        if not edits:
            # запасний варіант: клік у нижню частину + send_keys
            rect = win.rectangle()
            cx = rect.left + rect.width() // 2
            cy = rect.bottom - 60
            if PYAUTOGUI_AVAILABLE:
                pyautogui.click(cx, cy)
                time.sleep(0.1)
                safe_text = text.replace(" ", "{SPACE}")
                send_keys(safe_text, with_newlines=True, pause=0.01)
                if do_send:
                    send_keys("{ENTER}")
                return True
            return False

        def bottom_y(ed):
            try:
                r = ed.rectangle()
                return r.bottom
            except Exception:
                return -1

        edits.sort(key=bottom_y, reverse=True)
        edit = edits[0]

        try:
            edit.set_focus()
        except Exception:
            pass
        try:
            edit.click_input()
        except Exception:
            pass
        time.sleep(0.1)

        safe_text = text.replace(" ", "{SPACE}")
        send_keys(safe_text, with_newlines=True, pause=0.01)
        time.sleep(0.05)
        if do_send:
            send_keys("{ENTER}")
            time.sleep(0.05)
        return True
    except Exception as e:
        log_message(f"pywinauto друк: {e}")
        return False

# --- ctypes, щоб дістати PID з HWND (без win32api)
def _get_pid_from_hwnd(hwnd: int) -> int:
    user32 = ctypes.windll.user32
    pid = wt.DWORD()
    user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
    return pid.value

def _pyautogui_activate_win():
    """Шукаємо вікно WhatsApp через pyautogui, ігноруємо наше Tk-вікно/заголовок."""
    if not PYAUTOGUI_AVAILABLE:
        return None
    titles = ["WhatsApp Desktop", "WhatsApp", "WhatsApp Web", "WhatsApp Business"]
    for t in titles:
        try:
            wins = pyautogui.getWindowsWithTitle(t)
        except Exception:
            wins = []
        for w in wins:
            if not w:
                continue
            title = (w.title or "")
            # фільтр нашого вікна
            if APP_TITLE and APP_TITLE in title:
                continue
            if "Tk" in title and "TopLevel" in title:
                continue
            # PID/процес
            try:
                hwnd = int(getattr(w, "_hWnd", 0))
            except Exception:
                hwnd = 0
            if hwnd:
                try:
                    pid = _get_pid_from_hwnd(hwnd)
                    if pid == SELF_PID:
                        continue
                    if PSUTIL_AVAILABLE:
                        pname = psutil.Process(pid).name()
                        if "whatsapp" not in pname.lower():
                            continue
                except Exception:
                    pass
            return w
    return None

def _pyautogui_paste(text: str, do_send: bool, pre_ms: int, paste_delay_s: float, send_delay_s: float) -> bool:
    """Fallback: активуємо вікно, клікаємо в композер, Ctrl+V, Enter."""
    w = _pyautogui_activate_win()
    if not w:
        return False
    try:
        if w.isMinimized:
            w.restore()
            time.sleep(0.3)
        for _ in range(3):
            w.activate()
            time.sleep(0.2)
            aw = None
            try:
                aw = pyautogui.getActiveWindow()
            except Exception:
                pass
            if aw and "WhatsApp" in (aw.title or ""):
                break
            pyautogui.click(w.left + 80, w.top + 10)  # титулбар
            time.sleep(0.15)

        # клік у нижню середину (композер)
        pyautogui.press('esc')
        time.sleep(0.05)
        cx = w.left + w.width // 2
        cy = w.top + w.height - 60
        pyautogui.click(cx, cy)
        time.sleep(max(0.0, pre_ms/1000.0))

        if not PYPERCLIP_AVAILABLE:
            return False
        pyperclip.copy(text)
        time.sleep(0.15)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(max(0.05, paste_delay_s))
        if do_send:
            pyautogui.press('enter')
            time.sleep(max(0.05, send_delay_s))
        return True
    except Exception as e:
        log_message(f"pyautogui вставка: {e}")
        return False

def whatsapp_send(text: str, do_send=True, pre_ms=300, paste_delay_s=0.5, send_delay_s=0.3) -> bool:
    """UIA-друк (з {SPACE}) → fallback Ctrl+V. 3 спроби."""
    for attempt in range(1, 4):
        if PYWINAUTO_AVAILABLE and _pywinauto_focus_and_type(text, do_send):
            log_message(f"✅ Вставка через UIA (спроба {attempt})")
            return True
        ok = _pyautogui_paste(text, do_send, pre_ms, paste_delay_s, send_delay_s)
        if ok:
            log_message(f"✅ Вставка через Ctrl+V (спроба {attempt})")
            return True
        time.sleep(0.2 * attempt)
    return False

# ================== ВІДПРАВКА/ТАЙМЕР ==================
def do_send_report(text, pre_ms, paste_s, send_s, via_timer=False):
    # антифлуд: 15 хв
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
    """
    Фоновий цикл таймера.
    - Планування next_report_time робиться тільки тут.
    - На один target — лише ОДНЕ відправлення (fire_lock + last_fired_target).
    - Після спрацювання наступний target = get_next_slot(prev_target + 1 година).
    """
    global next_report_time, timer_active, last_fired_target
    with state_lock:
        timer_active = True
        if next_report_time is None:
            next_report_time = get_next_slot()  # первинна ініціалізація
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
            # атомарний захист від одночасних спрацювань
            if not fire_lock.acquire(blocking=False):
                time.sleep(0.1)
                continue
            try:
                log_message(f"⏰ ТАЙМЕР: {target.strftime('%H:%M:%S')} — відправляю автоматично.")
                with state_lock:
                    last_fired_target = target  # позначаємо, що цей target уже обробляється

                # зчитуємо GUI-параметри в головному треді → воркер як кнопка
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

                # одразу плануємо наступний target ВІД поточного target + 1 година
                with state_lock:
                    next_report_time = get_next_hour_slot_from_target(target)
                    log_message(f"📅 Наступна доповідь запланована на {next_report_time.strftime('%H:%M:%S')}")

            finally:
                fire_lock.release()

        time.sleep(0.2)

# ================== GUI ==================
root = tk.Tk()
root.title(APP_TITLE)
root.geometry("920x720")

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
        # не даємо стартувати другому треду
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

row = tk.Frame(actions); row.pack(pady=10)
tk.Button(row, text="Відправити зараз", command=send_now, font=("Arial", 9), bg="#2196F3", fg="white", width=17).pack(side=tk.LEFT, padx=3)
tk.Button(row, text="Тест вставлення", command=test_insert, font=("Arial", 9), bg="#FF9800", fg="white", width=17).pack(side=tk.LEFT, padx=3)
tk.Button(row, text="Діагностика", command=diagnose, font=("Arial", 9), bg="#9C27B0", fg="white", width=17).pack(side=tk.LEFT, padx=3)

# вкладка логів
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

# --- лог-помпа
def pump_logs():
    try:
        while True:
            line = log_q.get_nowait()
            log_text.insert(tk.END, line)
            log_text.see(tk.END)
    except queue.Empty:
        pass
    root.after(50, pump_logs)

# --- таймерний лейбл: лише показує (НЕ змінює next_report_time)
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

# ================== СТАРТ ==================
root.title(APP_TITLE)
log_message("🚀 Запуск. Рекомендовано: pip install pywinauto psutil")

# не плануємо нічого тут — планування робить лише таймер-тред;
# лейбл сам рахує відображення
root.after(0, pump_logs)
root.after(0, update_timer_label)
root.mainloop()
