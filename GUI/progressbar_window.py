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


# import requests
# import tkinter as tk
# from threading import Thread
# from PIL import Image, ImageTk
# from tkinter import ttk
#
#
# class PictureDownload(Thread):
#     def __init__(self, url):
#         super().__init__()
#
#         self.picture_file = None
#         self.url = url
#
#     def run(self):
#         """ download a picture and save it to a file """
#         # download the picture
#         response = requests.get(self.url)
#         picture_name = self.url.split('/')[-1]
#         picture_file = f'./assets/{picture_name}.jpg'
#
#         # save the picture to a file
#         with open(picture_file, 'wb') as f:
#             f.write(response.content)
#
#         self.picture_file = picture_file
#
#
# class App(tk.Tk):
#     def __init__(self, canvas_width, canvas_height):
#         super().__init__()
#         self.resizable(0, 0)
#         self.title('Image Viewer')
#
#         # Progress frame
#         self.progress_frame = ttk.Frame(self)
#
#         # configrue the grid to place the progress bar is at the center
#         self.progress_frame.columnconfigure(0, weight=1)
#         self.progress_frame.rowconfigure(0, weight=1)
#
#         # progressbar
#         self.pb = ttk.Progressbar(
#             self.progress_frame, orient=tk.HORIZONTAL, mode='indeterminate')
#         self.pb.grid(row=0, column=0, sticky=tk.EW, padx=10, pady=10)
#
#         # place the progress frame
#         self.progress_frame.grid(row=0, column=0, sticky=tk.NSEW)
#
#         # Picture frame
#         self.picture_frame = ttk.Frame(self)
#
#         # canvas width &amp; height
#         self.canvas_width = canvas_width
#         self.canvas_height = canvas_height
#
#         # canvas
#         self.canvas = tk.Canvas(
#             self.picture_frame,
#             width=self.canvas_width,
#             height=self.canvas_height)
#         self.canvas.grid(row=0, column=0)
#
#         self.picture_frame.grid(row=0, column=0)
#
#         # Button
#         btn = ttk.Button(self, text='Next Picture')
#         btn['command'] = self.handle_download
#         btn.grid(row=1, column=0)
#
#     def start_downloading(self):
#         self.progress_frame.tkraise()
#         self.pb.start(20)
#
#     def stop_downloading(self):
#         self.picture_frame.tkraise()
#         self.pb.stop()
#
#     def set_picture(self, file_path):
#         """ Set the picture to the canvas """
#         pil_img = Image.open(file_path)
#
#         # resize the picture
#         resized_img = pil_img.resize(
#             (self.canvas_width, self.canvas_height),
#             Image.ANTIALIAS)
#
#         self.img = ImageTk.PhotoImage(resized_img)
#
#         # set background image
#         self.bg = self.canvas.create_image(
#             0,
#             0,
#             anchor=tk.NW,
#             image=self.img)
#
#     def handle_download(self):
#         """ Download a random photo from unsplash """
#         self.start_downloading()
#
#         url = 'https://source.unsplash.com/random/640x480'
#         download_thread = PictureDownload(url)
#         download_thread.start()
#
#         self.monitor(download_thread)
#
#     def monitor(self, download_thread):
#         """ Monitor the download thread """
#         if download_thread.is_alive():
#             self.after(100, lambda: self.monitor(download_thread))
#         else:
#             self.stop_downloading()
#             self.set_picture(download_thread.picture_file)
#
#
# if __name__ == '__main__':
#     app = App(640, 480)
#     app.mainloop()

import tkinter as tk
from tkinter import ttk
import threading
import time


def foo(progress_var):
    for i in range(101):
        time.sleep(0.1)  # Замените это на вашу длительную операцию
        progress_var.set(i)
        root.update_idletasks()


def start_task():
    start_button.config(state=tk.DISABLED)  # Блокируем кнопку во время выполнения
    progress_var.set(0)

    thread = threading.Thread(target=foo, args=(progress_var,))
    thread.start()

    root.after(100, check_thread, thread)


def check_thread(thread):
    if thread.is_alive():
        root.after(100, check_thread, thread)
    else:
        progress_var.set(100)
        start_button.config(state=tk.NORMAL)  # Разблокируем кнопку после завершения


root = tk.Tk()
root.title("ProgressBar Example")

progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, mode="determinate", variable=progress_var)
progress_bar.pack(pady=10)

start_button = tk.Button(root, text="Start Task", command=start_task)
start_button.pack()

root.mainloop()