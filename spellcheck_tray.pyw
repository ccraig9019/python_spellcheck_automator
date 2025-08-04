import pyperclip
import win32com.client
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw
import threading
import ctypes
import pythoncom
import time
import pyautogui
from pywinauto import Application

def focus_word_window():
    try:
        app = Application(backend="uia").connect(path="WINWORD.EXE")
        # You might have multiple windows; pick the main one
        word_window = app.top_window()
        word_window.set_focus()
        time.sleep(0.5)  # Small pause to ensure focus
    except Exception as e:
        print("Could not focus Word window:", e)

def run_spellcheck():
    try:
        # Initialize COM in this thread
        pythoncom.CoInitialize()

        text = pyperclip.paste()

        if not text.strip():
            ctypes.windll.user32.MessageBoxW(0, "Clipboard is empty.", "Error", 0)
            return

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True

        doc = word.Documents.Add()
        word.Selection.TypeText(text)
        word.Selection.WholeStory()
        word.Selection.LanguageID = 1033  # US English
        
        focus_word_window()

        # Wait to ensure Word is focused
        time.sleep(1)

        # Simulate F7 press to open Editor
        pyautogui.press('f7')

        ctypes.windll.user32.MessageBoxW(0, "Spellcheck launched.", "Success", 0)

    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Error:\n{str(e)}", "Exception", 0)
    finally:
        pythoncom.CoUninitialize()


def run_in_thread():
    threading.Thread(target=run_spellcheck, daemon=True).start()

def create_icon_image():
    image = Image.new('RGB', (64, 64), "white")
    draw = ImageDraw.Draw(image)
    draw.text((12, 20), "ABC", fill="black")
    return image

menu = Menu(
    MenuItem("Run Spellcheck", run_in_thread),
    MenuItem("Quit", lambda icon, item: icon.stop())
)

icon = Icon("SpellcheckTray", icon=create_icon_image(), menu=menu)
icon.run()