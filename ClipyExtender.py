import pyperclip
import time
import pystray
from pystray import MenuItem as Item, Icon, Menu
from PIL import Image, ImageDraw
import threading
import json
import os
import sys
import html
import qrcode
import pygame
import winreg
import requests

# === 基本設定 ===
CONFIG_FILE = "config.json"
monitoring = True

# === PyInstaller の _MEIPASS 対応 ===
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

wav_file_path = os.path.join(base_path, "notification_sound.wav")

# === pygame 初期化 ===
pygame.mixer.init()

# 設定をロード
def load_config():
    default_config = {
        "interval": 0.5  # デフォルトの監視間隔
    }
    if not os.path.exists(CONFIG_FILE):
        save_config(default_config)
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

# 設定を保存
def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

config = load_config()
monitoring_interval = config["interval"]

# === NGワードのロード ===
def load_ng_words():
    ng_words_file = "ng_words.json"
    if not os.path.exists(ng_words_file):
        default_ng_words = ["password is 1234"]
        save_ng_words(default_ng_words)
    with open(ng_words_file, "r", encoding="utf-8") as f:
        ng_words = json.load(f)
        print(f"NG Words Loaded: {ng_words}")  # ここでNGワードの内容を確認
        return ng_words

# === NGワードを保存 ===
def save_ng_words(ng_words):
    with open("ng_words.json", "w", encoding="utf-8") as f:
        json.dump(ng_words, f, indent=4, ensure_ascii=False)

# === クリップボードのNGワードフィルタリング ===
def filter_ng_words(text, ng_words):
    for word in ng_words:
        if word in text:
            text = text.replace(word, "*****")
            pyperclip.copy(text)
            break
    
    return text

# === クリップボード監視スレッド ===
def modify_clipboard():
    global monitoring, monitoring_interval
    url_mapping = load_url_mapping()
    ng_words = load_ng_words()
    
    while True:
        if monitoring:
            text = pyperclip.paste()
            text = convert_to_plain_text(text)
            text = filter_ng_words(text, ng_words)
            
            # 短縮URL展開
            if any(short in text for short in ["bit.ly", "t.co", "tinyurl.com"]):
                expanded_url = expand_short_url(text)
                if expanded_url != text:
                    pyperclip.copy(expanded_url)
                    play_sound()

            for old_url, new_url in url_mapping.items():
                if text.startswith(old_url):
                    new_text = text.replace(old_url, new_url, 1)
                    if new_text != text:
                        pyperclip.copy(new_text)
                        play_sound()
                    break
        time.sleep(monitoring_interval)


# === 音を再生 ===
def play_sound():
    if os.path.exists(wav_file_path):
        pygame.mixer.music.load(wav_file_path)
        pygame.mixer.music.play()
    else:
        print(f"⚠ WAVファイルが見つからない: {wav_file_path}")

# === 監視間隔変更 ===
def change_interval(icon, item, new_interval):
    global monitoring_interval
    monitoring_interval = new_interval
    config["interval"] = new_interval
    save_config(config)
    update_menu(icon)

# === メニュー更新 ===
def update_menu(icon):
    monitor_text = "一時停止" if monitoring else "再開"
    startup_text = "スタートアップに登録" if not is_startup_enabled() else "スタートアップから解除"
    
    icon.menu = Menu(
        Item(monitor_text, toggle_monitoring),
        Item(f"現在の監視間隔: {monitoring_interval}秒", lambda icon, item: None),
        Item("監視間隔を0.1秒に変更", lambda icon, item: change_interval(icon, item, 0.1)),
        Item("監視間隔を0.5秒に変更", lambda icon, item: change_interval(icon, item, 0.5)),
        Item("監視間隔を1.0秒に変更", lambda icon, item: change_interval(icon, item, 1.0)),
        Item("QRコード生成", create_qr_from_clipboard),
        Item(startup_text, toggle_startup),
        Item("終了", exit_program)
    )
    icon.update_menu()

# === スタートアップ登録・解除 ===
def toggle_startup(icon, item):
    if is_startup_enabled():
        remove_from_startup()
    else:
        add_to_startup()
    update_menu(icon)

def is_startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as reg_key:
            winreg.QueryValueEx(reg_key, "ClipboardWatcher")
        return True
    except FileNotFoundError:
        return False

def add_to_startup():
    script_path = os.path.abspath(sys.argv[0])
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as reg_key:
        winreg.SetValueEx(reg_key, "ClipboardWatcher", 0, winreg.REG_SZ, script_path)

def remove_from_startup():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE) as reg_key:
            winreg.DeleteValue(reg_key, "ClipboardWatcher")
    except FileNotFoundError:
        pass

# === クリップボード関連関数 ===
def convert_to_plain_text(text):
    return html.unescape(text)

def load_url_mapping():
    mapping_file = "url_mapping.json"
    if not os.path.exists(mapping_file):
        default_mapping = {"https://x.com/": "https://fixvx.com/", "https://twitter.com/": "https://fixvx.com/"}
        save_url_mapping(default_mapping)
    with open(mapping_file, "r", encoding="utf-8") as f:
        return json.load(f)

def save_url_mapping(mapping):
    with open("url_mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=4, ensure_ascii=False)

# === QRコード生成 ===
def create_qr_from_clipboard(icon, item):
    text = pyperclip.paste()
    text = convert_to_plain_text(text)
    if text.startswith("https://") or text.startswith("http://"):
        qrcode.make(text).show()

# === アイコン作成 ===
def create_icon():
    image = Image.new("RGB", (64, 64), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    draw.ellipse((8, 8, 56, 56), fill=(0, 0, 0))
    return image

# === shortlink opener ===
def expand_short_url(url):
    try:
        response = requests.head(url, allow_redirects=True, timeout=5)
        return response.url
    except requests.RequestException:
        return url

# === 監視ON/OFF ===
def toggle_monitoring(icon, item):
    global monitoring
    monitoring = not monitoring
    update_menu(icon)

# === 終了 ===
def exit_program(icon, item):
    icon.stop()

# === メイン処理 ===
icon = Icon("clipboard_watcher", create_icon())
update_menu(icon)

clipboard_thread = threading.Thread(target=modify_clipboard, daemon=True)
clipboard_thread.start()

icon.run()
