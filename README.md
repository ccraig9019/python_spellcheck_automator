This is a simple Python script created with the aid of AI. 
In my work as a freelance translator, I use web-based tools which do not have a built-in spellchecker. When I complete a piece of work, the recommended course of action is to copy the contents of the work to the clipboard and paste them into a Word document, making sure to change the proofreading language to US English. Word's spellchecker and grammar checker are then used to find any errors, and corrections are made directly in the web-based tool
This process, although simple, can become time consuming when repeated several times a day. I built this script to automate the process of opening Word, pasting in the contents of the clipboard, setting the language to US English, then opening the editor.
This script uses several Python libraries, including:
- pyperclip, for access to the clipboard
- pystray, to create an icon accessible in the system tray
- pywinauto, to automate the use of keyboard shortcuts
