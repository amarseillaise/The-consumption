# from tkinter import ttk
# import tkinter as tk
#
#
# class ProgressBarWindow:
#     def __init__(self):
#         self.pb_window = tk.Tk()
#         self.pb_window.geometry('300x60')
#         self.pb_window.resizable(False, False)
#         self.pb_window.title('Рассчитываем...')
#
#         # progressbar
#         self.pb = ttk.Progressbar(
#             self.pb_window,
#             orient='horizontal',
#             mode='determinate',
#             maximum=100,
#             value=0,
#             length=280
#         )
#         self.pb.grid(column=0,
#                      row=0,
#                      columnspan=2,
#                      padx=10,
#                      pady=8
#                      )
#
#         # label
#         self.value_label = ttk.Label(self.pb_window,
#                                      text=f"{self.pb['value']}%"
#                                      )
#         self.value_label.grid(column=0,
#                               row=1,
#                               columnspan=2,
#                               )
#
#         self.pb_window.mainloop()
#
#     def update(self, value: int):
#         self.pb['value'] += value
#         self.value_label['text'] = f"{self.pb['value']}%"


import tkinter as tk
from tkinter import ttk
from tqdm import tqdm
from joblib import Parallel, delayed
from threading import Thread

n = 0


def fun(_):
    global n
    n += 1
    return _ ** 2


def main():
    def parallel():
        nonlocal result
        result = Parallel(n_jobs=-1, backend='threading')(delayed(fun)(_) for _ in range(1024 ** 2))

    result = None
    progress_bar['maximum'] = 1024 ** 2  # number of items that loops in Parallel
    process = Thread(target=parallel, daemon=True)
    process.start()
    """
    update progress bar
    """
    while progress_bar['value'] < 1024 ** 2:
        progress_bar['value'] = n
        root.update_idletasks()
        root.update()  # prevent freezing
    process.join()

    print(result)


root = tk.Tk()
progress_bar = ttk.Progressbar(root, mode='determinate')
progress_bar.pack()
button = tk.Button(root, text='start', command=main)
button.pack()
root.mainloop()