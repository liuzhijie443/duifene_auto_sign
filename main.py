import re
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import requests
import urllib3
from bs4 import BeautifulSoup
# from pyzbar.pyzbar import decode

host = "https://www.duifene.com"
UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 MicroMessenger/8.0.40(0x1800282a) NetType/WIFI Language/zh_CN'
# 近代史
course_id = None
class_id = None
# 测试
# course_id = "239781"
# class_id = "400636"
class_json = []
seconds = 5
job_id = ""
username = ""
password = ""
stu_id = ""

urllib3.disable_warnings()
x = requests.Session()
x.headers['User-Agent'] = UA
# x.proxies = {"https": "127.0.0.1:8888"}
x.verify = False
r = x.get(url=host)


# def qrcode(base64_image):
#     # 将base64编码的图片转换为二进制数据
#     image_data = base64.b64decode(base64_image)
#     # 读取图片数据并转换为PIL的Image对象
#     image = Image.open(BytesIO(image_data))
#     # 使用pyzbar的decode函数扫描二维码
#     decoded_image = decode(image)
#     for barcode in decoded_image:
#         text = barcode.data.decode("utf-8")  # 打印扫描出的二维码内容
#         return text


def get_user_id():
    global stu_id
    r1 = x.get(url=host + "/_UserCenter/MB/index.aspx")
    if r1.status_code == 200:
        soup = BeautifulSoup(r1.text, "lxml")
        code = soup.find(id="hidUID").get("value")
        stu_id = code


def sign(code):
    status = False
    if len(code) == 4:
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "https://www.duifene.com/_CheckIn/MB/CheckInStudent.aspx?moduleid=16&pasd="
        }
        params = f"action=studentcheckin&studentid={stu_id}&checkincode={code}"
        r1 = x.post(
            url=host + "/_CheckIn/CheckIn.ashx", data=params, headers=headers)
        if r1.status_code == 200:
            msg = r1.json()["msgbox"]
            text_box.insert(tk.END, f"\n\n{msg}\n")
            if msg == "签到成功！":
                text_box.insert(tk.END, f"\n")
                text_box.insert(tk.END, f"\n")
                text_box.insert(tk.END, f"\n")
                status = True
    else:
        r1 = x.get(url=host + "/_CheckIn/MB/QrCodeCheckOK.aspx?state=" + code)
        if r1.status_code == 200:
            soup = BeautifulSoup(r1.text, "lxml")
            msg = soup.find(id="DivOK").get_text()
            if "签到成功" in msg:
                text_box.insert(tk.END, f"\n\n{msg}\n")
                text_box.insert(tk.END, f"\n")
                text_box.insert(tk.END, f"\n")
                text_box.insert(tk.END, f"\n")
                status = True
            else:
                text_box.insert(tk.END, f"非微信链接登录，二维码无法签到")
    return status


def get_code():
    global job_id
    sign_status = False

    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    text_box.insert(tk.END, f"\n\n持续监控：{current_time}")  # 插入当前时间
    text_box.see(tk.END)  # 滚动到最后一行

    r1 = x.get(
        url=host + f"/_CheckIn/MB/TeachCheckIn.aspx?classid={class_id}&temps=0&checktype=1&isrefresh=0&timeinterval=0&roomid=0&match=")
    if r1.status_code == 200:
        if "HFCheckCodeKey" in r1.text:
            soup = BeautifulSoup(r1.text, "lxml")
            code = soup.find(id="HFCheckCodeKey").get("value")
            HFSeconds = soup.find(id="HFSeconds").get("value")
            HFChecktype = soup.find(id="HFChecktype").get("value")
            HFCheckInID = soup.find(id="HFCheckInID").get("value")
            # 数字签到
            if HFChecktype == '1':
                text_box.insert(tk.END, f"\t未到签到时间\t倒计时：{HFSeconds}秒\t签到码：{code}")
                if code is not None and int(HFSeconds) <= int(seconds):
                    text_box.insert(tk.END, f"\n\n开始签到\t签到码：{code}")
                    get_user_id()
                    sign_status = sign(code)
            # 二维码签到
            elif HFChecktype == '2':
                if HFCheckInID is not None:
                    text_box.insert(tk.END, f"\n\n开始签到\t二维码签到")
                    get_user_id()
                    sign_status = sign(HFCheckInID)
    if sign_status is False:
        job_id = root.after(1000, get_code)
    else:
        text_box.insert(tk.END, f"\n")
        messagebox.showinfo("签到状态", "签到成功")


def enter():
    if course_id is None:
        messagebox.showerror("", "请先登录")
        return
    headers = {
        "Referer": "https://www.duifene.com/_UserCenter/MB/index.aspx"
    }
    r1 = x.get(url=host + "/_UserCenter/MB/Module.aspx?data=" + course_id, headers=headers)
    if r1.status_code == 200:
        if course_id in r1.text:
            soup = BeautifulSoup(r1.text, "lxml")
            CourseName = soup.find(id="CourseName").text
            text_box.insert(tk.END, f"\n开始监听【{CourseName}】签到活动")
            get_code()


def get_class_json():
    global class_json, course_id,class_id
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/_UserCenter/PC/CenterStudent.aspx"
    }
    params = "action=getstudentcourse&classtypeid=2"
    r1 = x.post(
        url=host + "/_UserCenter/CourseInfo.ashx", data=params, headers=headers)
    if r1.status_code == 200:
        j = r1.json()
        if j is not None:
            try:
                msg = j["msgbox"]
                messagebox.showerror("", f"{msg} 可能是登录链接有误，请重新复制。")
            except Exception as e:
                messagebox.showinfo("提示", "登录成功")
                class_list = []
                class_json = j
                for i in j:
                    class_list.append(i["CourseName"])
                combo['values'] = tuple(class_list)
                combo.set(class_list[0])
                course_id = 0
                class_id = j[0]["TClassID"]
                course_id = j[0]["CourseID"]


def login_link():
    link = link_entry.get()
    code = re.search(r"(?<=code=)\S{32}", link)
    if code is not None:
        code = code[0]
        x.get(url=host + f"/P.aspx?authtype=1&code={code}&state=1")
        get_class_json()
    else:
        messagebox.showerror("error", "链接有误")


def on_combo_change(event):
    global class_id, course_id
    className = combo_var.get()
    for i in class_json:
        if i["CourseName"] == className:
            class_id = i["TClassID"]
            course_id = i["CourseID"]


def login():
    global username, password, seconds
    username = username_entry.get()
    password = password_entry.get()
    seconds = seconds_entry.get()
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/AppGate.aspx"
    }
    params = f'action=loginmb&loginname={username}&password={password}'
    r1 = x.post(url=host + "/AppCode/LoginInfo.ashx", data=params, headers=headers)
    if r1.status_code == 200:
        text_box.delete("1.0", "end")
        msg = r1.json()["msgbox"]
        messagebox.showinfo("提示", msg)
        text_box.insert(tk.END, f"\n{msg}\n")
        if msg == "登录成功":
            get_class_json()
    else:
        messagebox.showerror("登录失败", "错误提示")


def select_tab(event):
    tab_id = tab_control.index(tab_control.select())
    if tab_id == 0:
        text = '''
        1、打开电脑端微信，复制如下链接到文件传输助手并发送\n
        【https://open.weixin.qq.com/connect/oauth2/authorize?appid=wx1b5650884f657981&redirect_uri=https://www.duifene.com/_FileManage/PdfView.aspx?file=https%3A%2F%2Ffs.duifene.com%2Fres%2Fr2%2Fu6106199%2F%E5%AF%B9%E5%88%86%E6%98%93%E7%99%BB%E5%BD%95_876c9d439ca68ead389c.pdf&response_type=code&scope=snsapi_userinfo&connect_redirect=1#wechat_redirect】\n\n
        2、点击进入链接，点击微信浏览器窗口右上角三个点，点击复制链接，并把微信链接粘贴到左侧输入框。
        '''
        text_box.insert(tk.END, text)
        frame_left_2.pack_forget()
        frame_left.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))
    elif tab_id == 1:
        text_box.delete("1.0", "end")
        frame_left.pack_forget()
        frame_left_2.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))


root = tk.Tk()
root.title("GUI测试")
root.resizable(False, False)
# root.configure(bg='lightgrey')

tab_control = ttk.Notebook(root)
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)
tab_control.add(tab1, text="微信链接登录")
tab_control.add(tab2, text="账号密码登录")
tab_control.pack(fill=tk.BOTH, side=tk.LEFT)
# 当选项卡被选中时，调用select_tab函数
tab_control.bind("<<NotebookTabChanged>>", select_tab)

frame_left = tk.Frame(tab_control)
frame_left.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))

text4 = tk.Label(frame_left, text="支持二维码和签到码", font=('宋体', 10))
text4.pack(pady=5)
text = tk.Label(frame_left, text="查看右侧说明进行登录", font=('宋体', 10))
text.pack(pady=5)
text2 = tk.Label(frame_left, text="链接", font=('宋体', 10))
text2.pack(pady=5)
link_entry = tk.Entry(frame_left, font=('宋体', 12))
link_entry.pack(pady=5, padx=10)
login_button = tk.Button(frame_left, text="登录", command=login_link, font=('宋体', 14))
login_button.pack(pady=5)

frame_left_2 = tk.Frame(tab_control)
text3 = tk.Label(frame_left_2, text="不支持二维码签到", font=('宋体', 10))
text3.pack(pady=5)

username_label = tk.Label(frame_left_2, text="账号", font=('宋体', 15))
username_label.pack(padx=10)
username_entry = tk.Entry(frame_left_2, font=('宋体', 12))
username_entry.pack(padx=10)

password_label = tk.Label(frame_left_2, text="密码", font=('宋体', 15))
password_label.pack(padx=10)
password_entry = tk.Entry(frame_left_2, show="*", font=('宋体', 12))
password_entry.pack(padx=10)

username_label = tk.Label(frame_left_2, text="剩余倒计时X秒后签到", font=('宋体', 10))
username_label.pack(pady=5)
seconds_entry = tk.Entry(frame_left_2, font=('宋体', 12), width=5)
seconds_entry.insert(0, "5")
seconds_entry.pack(pady=5)

login_button = tk.Button(frame_left_2, text="登录", command=login, font=('宋体', 14))
login_button.pack(pady=5)

frame_mid = tk.Frame(root)
frame_mid.pack(side=tk.TOP)
f2lable = tk.Label(frame_mid, text="选择课程")
f2lable.pack(side=tk.TOP, fill=tk.BOTH, pady=(10, 0))
combo_var = tk.StringVar()
combo = ttk.Combobox(frame_mid, textvariable=combo_var, state="readonly")
combo.bind("<<ComboboxSelected>>", on_combo_change)
combo.pack(side=tk.LEFT, )
f2btn = tk.Button(frame_mid, text="开始监听签到", command=enter)
f2btn.pack(side=tk.RIGHT, padx=10, pady=10)

frame_right = tk.Frame(root)
frame_right.pack(side=tk.RIGHT)

text_box = tk.Text(frame_right, width=90, height=20, font=('宋体', 9))
text_box.pack(pady=(0, 10), padx=(0, 10))

root.mainloop()
