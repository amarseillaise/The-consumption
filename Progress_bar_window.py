import tkinter as tk
from tkinter import ttk


class ProgressBar:
    
    def __int__(self):

        self.root = tk.Tk()
        self.root.title("Прогресс")
        self.root.geometry("300x150")
        self.root.resizable(False, False)

        frame = tk.Frame(self.root, padx=10, pady=10)
        frame.pack(expand=True)

        progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        progress.pack(pady=10)

        status_label = tk.Label(frame, text="Progress: 0%")
        status_label.pack()

        self.root.mainloop()
