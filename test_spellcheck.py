import win32com.client
import pyperclip

# Get text from clipboard
text = pyperclip.paste()

# Launch Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = True

# Create a new document and insert the text
doc = word.Documents.Add()
word.Selection.TypeText(text)

# Select all text and set US English
word.Selection.WholeStory()
word.Selection.LanguageID = 1033

# Start spellcheck
word.ActiveDocument.CheckSpelling()