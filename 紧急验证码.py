import datetime
import tkinter as tk
import hashlib

# 获取本地时间并格式化为"年-月-日-时-分"
def get_formatted_local_time():
    now = datetime.datetime.now()
    formatted_time = now.strftime("%Y%m%d%H%M")
    return formatted_time

def get_password_hash(user_input):
    try:
        number = int(user_input)
        number = str(number)
    # 计算 SHA-256 哈希值
        sha256_hash = hashlib.sha256(number.encode('utf-8')).hexdigest()
        #print(f"输入整数的SHA-256哈希值: {sha256_hash}")
        return sha256_hash
    except ValueError:
        print("输入无效的整数")

def get_password_md5(user_input):
    try:
        number = int(user_input)
        number = str(number)
    #计算md5
        md5 = hashlib.md5(number.encode('utf-8')).hexdigest()
        #print(f"输入整数的MD5哈希值: {md5_hash}")
        return md5
    except ValueError:
        print("输入无效的整数")

def zhongqi_yzm():
    hash256 = get_password_hash(get_formatted_local_time())
    copyword_entry.delete(0,tk.END)
    copyword_entry.insert(0,hash256)

def duanqi_yzm():
    md5 = get_password_md5(get_formatted_local_time())
    copyword_entry.delete(0,tk.END)
    copyword_entry.insert(0,md5)

#关闭窗口
def close_UI():
    root.destroy()

def copypassword():
    text_copy = copyword_entry.get()
    #清空粘贴板
    root.clipboard_clear()
    #写入粘贴板
    root.clipboard_append(text_copy)
#-------------------------------------------------------------

# 创建GUI窗口
root = tk.Tk()
root.title("excal解密器秘钥生成")
root.geometry("300x300")  # 设置窗口大小为？ x ？像素

# 创建一个Frame来容纳控件，并使其在窗口中居中
frame = tk.Frame(root)
frame.place(relx=0.5, rely=0.5, anchor="center")
# 创建一个新的输入框用于用户复制
copyword_label = tk.Label(frame, text="请复制")
copyword_label.pack()
copyword_entry = tk.Entry(frame, show="")
copyword_entry.pack(pady=5)

# 创建一个按钮，获取短期验证码
duanqi_button = tk.Button(frame, text="获取短期验证码", command=duanqi_yzm)
duanqi_button.pack(pady=3,ipadx=10,ipady=3,fill="both")

# 创建一个按钮，获取中期验证码
zhongqi_button = tk.Button(frame, text="获取中期验证码", command=zhongqi_yzm)
zhongqi_button.pack(pady=3,ipadx=10,ipady=3,fill="both")

# 创建一个按钮，复制
copy_button = tk.Button(frame, text="复制秘钥", command=copypassword)
copy_button.pack(pady=3,ipadx=10,ipady=3,fill="both")

# 创建一个按钮，关闭程序
close_button = tk.Button(frame, text="关闭", command=close_UI)
close_button.pack(pady=3,ipadx=10,ipady=3,fill="both")

# 运行GUI程序
root.mainloop()