from docx import Document
from collections import Counter
import string
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import scrolledtext

def read_docx(file):
    doc = Document(file)
    result = []
    for paragraph in doc.paragraphs:
        text = paragraph.text
        text = text.translate(str.maketrans('', '', string.punctuation)) # remove punctuation
        words = text.split()
        result.extend(words)
    return result

def print_unique_words(words, case_sensitive):
    if not case_sensitive:
        words = [word.lower() for word in words]
    word_counts = Counter(words)
    unique_words = [word for word, count in word_counts.items() if count == 1]
    return unique_words

def open_file():
    filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
    file_label.config(text=f"File Chosen: {filename}")
    words = read_docx(filename)
    unique_words = print_unique_words(words, case_sensitive_var.get())
    output_text.delete(1.0, tk.END) # Clear the text box before inserting new text
    for word in unique_words:
        output_text.insert(tk.END, f"{word}\n")

# Create the main window
window = tk.Tk()

# Create a label, button, and text box
file_label = tk.Label(window, text="No file chosen")
open_file_button = tk.Button(window, text="Open File", command=open_file)
output_text = scrolledtext.ScrolledText(window)
case_sensitive_var = tk.BooleanVar() # Boolean variable to hold the state of the checkbox
case_sensitive_check = tk.Checkbutton(window, text="Case Sensitive", variable=case_sensitive_var)

# Place the widgets in the window
file_label.pack(pady=10)
open_file_button.pack(pady=10)
case_sensitive_check.pack(pady=10)
output_text.pack(pady=10)

# Run the event loop
window.mainloop()
