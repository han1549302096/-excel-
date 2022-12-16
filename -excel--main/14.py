import tkinter as tk
from tkinter import ttk
from tkinter import messagebox


class LoginForm(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.name_label = tk.Label(self, text='用户名:')
        self.name_label.place(x=20, y=20)
        self.name_entry = tk.Entry(self, width=20)
        self.name_entry.place(x=65, y=20)
        self.pwd_label = tk.Label(self, text='密码:')
        self.pwd_label.place(x=20, y=50)
        self.pwd_entry = tk.Entry(self, width=20, show='*')
        self.pwd_entry.place(x=65, y=50)
        self.submit_button = tk.Button(self, text='登录', width=10, command=self.submit)
        self.submit_button.place(x=65, y=80)
        self.quit_button = tk.Button(self, text='退出', width=10, command=self.quit)
        self.quit_button.place(x=130, y=80)

    def submit(self):
        username = self.name_entry.get()
        password = self.pwd_entry.get()
        if username == 'admin' and password == '123456':
            messagebox.showinfo(title='登录成功', message='欢迎管理员')
        else:
            messagebox.showerror(title='登录失败', message='用户名或密码错误')


if __name__ == '__main__':
    root = tk.Tk()
    root.title('登录界面')
    root.geometry('250x120')
    app = LoginForm(master=root)
    root.mainloop()
