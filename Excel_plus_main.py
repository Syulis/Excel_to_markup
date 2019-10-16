import tkinter as tk
from tkinter import filedialog
import json
import re
import pandas as pd

#データファイル,フォルダ
data_json = "Excel_plus_data.json"
Excel_file = "Excel_plus"

ex = open(data_json, 'r')
Excel_data = json.load(ex)
ex.close()

class ui_window(tk.Frame):

    def __init__(self, master):
        tk.Frame.__init__(self, master)

        global Excel_data

        #window情報
        self.master.title("Excel Plus")
        self.master.geometry("445x420")

        #タブ
        menu = tk.Menu(self.master)
        self.master.config(menu=menu)

        #ファイルタブ
        menu_file = tk.Menu(self.master)
        menu.add_cascade(label="File", menu=menu_file)
        menu_file.add_command(label="Open", command=self.file_open, activebackground="blue")

        #共通変数
        self.open_file = None

        #テンプレート入力
        self.template_box = tk.Entry(width=47)
        self.template_box.place(x=10, y=21.5)

        #クリアボタン
        clear_button = tk.Button(text="All Clear")
        clear_button.place(x=200, y=52.5)
        clear_button.bind('<1>', self.all_clear)

        #実行ボタン
        execution_button = tk.Button(text="Execution")
        execution_button.place(x=270, y=52.5)
        execution_button.bind('<1>', self.execution)

        #出力箱
        self.text_box = tk.Text(width=60)
        self.text_box.place(x=10, y=90)

        #コピーボタン
        copy_button = tk.Button(text="Copy")
        copy_button.place(x=347.5, y=52.5)
        copy_button.bind('<1>', self.copy)

        #ログ箱
        self.log_box = tk.Entry(width=15, font=("Meiryo", "10"))
        self.log_box.place(x=307.5, y=20)
        self.log_box.delete(0, tk.END)
        self.log_box.insert(tk.END, "Log")

        #終了ボタン
        self.end_button = tk.Button(text="End")
        self.end_button.place(x=400, y=52.5)
        self.end_button.bind('<1>', self.end)

        # ファイル読み込み
        try:
            self.open_file = pd.read_excel(Excel_data["stat"]["open_file"], index_col=None, header=None)
            self.put_log("green", "Loaded")
        except:
            self.file_open()

    #ファイルオープン
    def file_open(self):
        fld = filedialog.askopenfilenames(initialdir="./Excel_File")
        Excel_data["stat"]["open_file"] = fld[0]
        self.open_file = pd.read_excel(Excel_data["stat"]["open_file"], index_col=None, header=None)
        self.put_log("green", "Loaded")

    #実行
    def execution(self, event):
        word = "%"
        self.text_box.delete('0.0', tk.END)
        after_text = self.template_box.get()
        loop_number = after_text.count(word)
        column = 0
        if loop_number == 0:
            self.put_log("red", "Not Found Letter to Change")
            return
        while True:
            try:
                for i in range(loop_number):
                    number = after_text[after_text.find(word) + 1]
                    text = str(self.open_file.iat[column, int(number) - 1])
                    if text == "nan":
                        text = "&nbsp;"
                    after_text = re.sub(word + number, text, after_text)

            except ValueError:
                self.text_box.insert(tk.END, after_text + "\n")
                after_text = self.template_box.get()
                column += 1
                pass

            except IndexError:
                if self.text_box.get('0.0', tk.END) == "\n":
                    self.put_log("red", "No Found Data")
                else:
                    self.put_log("blue", "Done")
                break

    #クリア
    def all_clear(self, event):
        self.template_box.delete(0, tk.END)
        self.text_box.delete('0.0', tk.END)
        self.put_log("blue", "All Clear")

    #コピー
    def copy(self, event):
        self.master.clipboard_clear()
        self.master.clipboard_append(self.text_box.get('1.0', tk.END).strip())
        self.put_log("green", "Copied")

    #ログ
    def put_log(self, color, text):
        self.log_box.delete(0, tk.END)
        self.log_box.configure(fg=color)
        self.log_box.insert(tk.END, text)

    #終了
    def end(self, event):
        with open(data_json, 'w') as f:
            json.dump(Excel_data, f, indent=3)
        self.master.destroy()

if __name__ == '__main__':
    f = ui_window(None)
    f.pack()
    f.mainloop()
