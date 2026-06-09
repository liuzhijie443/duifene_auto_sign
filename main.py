import configparser
import os.path
import re
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import requests
import urllib3
from bs4 import BeautifulSoup
import random


class Course:
    id = '0'
    class_id = '0'
    flag = True
    check_list = []
    class_list = []


def on_combo_change(event):
    className = combo_var.get()
    for i in Course.class_list:
        if i["CourseName"] == className:
            Course.id = i["CourseID"]
            Course.class_id = i["TClassID"]


def save_cookie(_x):
    config['INFO'] = {
        'cookie': _x.request.headers.get("cookie")
    }
    with open(filename, 'w') as f:
        config.write(f)


def login_link():
    link = link_entry.get()
    code = re.search(r"(?<=code=)\S{32}", link)
    if code is not None:
        x.cookies.clear()
        code = code[0]
        _r = x.get(url=host + f"/P.aspx?authtype=1&code={code}&state=1")
        get_class_list()
        save_cookie(_r)
    else:
        messagebox.showerror("error", "链接有误")


def get_remaining_seconds(limit_date_str):
    """计算距离签到截止时间还剩多少秒，如果已过期则返回负数"""
    try:
        limit_time = datetime.strptime(limit_date_str, "%Y/%m/%d %H:%M:%S")
        now = datetime.now()
        delta = (limit_time - now).total_seconds()
        return int(delta)
    except Exception:
        return 0


def login():
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/AppGate.aspx"
    }
    params = f'action=loginmb&loginname={username.get()}&password={password.get()}'
    x.cookies.clear()
    x.get(host)
    _r = x.post(url=host + "/AppCode/LoginInfo.ashx", data=params, headers=headers)
    if _r.status_code == 200:
        text_box.delete("1.0", "end")
        msg = _r.json()["msgbox"]
        text_box.insert(tk.END, f"\n{msg}\n")
        if msg == "登录成功":
            get_class_list()
            save_cookie(_r)
    else:
        messagebox.showerror("错误提示", "登录失败")


def select_tab(event):
    tab_id = tab_control.index(tab_control.select())
    text_box.delete("1.0", "end")
    if tab_id == 0:
        text = '''
        1、打开电脑端微信，复制如下链接到文件传输助手并发送\n
        【https://open.weixin.qq.com/connect/oauth2/authorize?appid=wx1b5650884f657981&redirect_uri=https://www.duifene.com/_FileManage/PdfView.aspx?file=https%3A%2F%2Ffs.duifene.com%2Fres%2Fr2%2Fu6106199%2F%E5%AF%B9%E5%88%86%E6%98%93%E7%99%BB%E5%BD%95_876c9d439ca68ead389c.pdf&response_type=code&scope=snsapi_userinfo&connect_redirect=1#wechat_redirect】\n\n
        2、点击进入链接，点击微信浏览器窗口右上角三个点，点击复制链接，并把微信链接粘贴到左侧输入框。\n
        '''
        text_box.insert(tk.END, text)
        tab_frame2.pack_forget()
        tab_frame1.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))
    elif tab_id == 1:
        tab_frame1.pack_forget()
        tab_frame2.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))


def get_user_id():
    _r = x.get(url=host + "/_UserCenter/MB/index.aspx")
    if _r.status_code == 200:
        soup = BeautifulSoup(_r.text, "lxml")
        stu_id = soup.find(id="hidUID").get("value")
        return stu_id


def sign(sign_code):
    # 签到码
    if len(sign_code) == 4:
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "https://www.duifene.com/_CheckIn/MB/CheckInStudent.aspx?moduleid=16&pasd="
        }
        params = f"action=studentcheckin&studentid={get_user_id()}&checkincode={sign_code}"
        _r = x.post(
            url=host + "/_CheckIn/CheckIn.ashx", data=params, headers=headers)
        if _r.status_code == 200:
            msg = _r.json()["msgbox"]
            text_box.insert(tk.END, f"\t{msg}\n\n")
            if msg == "签到成功！":
                return True
    # 二维码
    else:
        _r = x.get(url=host + "/_CheckIn/MB/QrCodeCheckOK.aspx?state=" + sign_code)
        if _r.status_code == 200:
            soup = BeautifulSoup(_r.text, "lxml")
            msg = soup.find(id="DivOK").get_text()
            if "签到成功" in msg:
                text_box.insert(tk.END, f"\t{msg}\n\n")
            else:
                text_box.insert(tk.END, f"\t非微信链接登录，二维码无法签到\n\n")
            return True


def sign_location(longitude, latitude):
    longitude = str(round(float(longitude) + random.uniform(-0.000089, 0.000089), 8))
    latitude = str(round(float(latitude) + random.uniform(-0.000089, 0.000089), 8))

    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/_CheckIn/MB/CheckInStudent.aspx?moduleid=16&pasd="
    }
    params = f"action=signin&sid={get_user_id()}&longitude={longitude}&latitude={latitude}"
    _r = x.post(
        url=host + "/_CheckIn/CheckInRoomHandler.ashx", data=params, headers=headers)
    if _r.status_code == 200:
        msg = _r.json()["msgbox"]
        text_box.insert(tk.END, f"\t{msg}\n\n")
        if msg == "签到成功！":
            return True


def watching_sign():
    print("watching_sign called", datetime.now())
    is_login()

    line_count = int(text_box.index('end-1c').split('.')[0])
    text_box.delete(f"{line_count}.0", f"{line_count}.end")
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    text_box.insert(tk.END, f"持续监控：{current_time}")
    text_box.see(tk.END)

    # 获取学生ID
    student_id = get_user_id()
    if not student_id:
        text_box.insert(tk.END, "\n获取学生ID失败，请重新登录\n")
        if Course.flag:
            root.after(5000, watching_sign)
        return

    # 构造 POST 请求
    post_data = {
        "action": "getstudentinlogbyday",
        "classid": Course.class_id,
        "studentid": student_id
    }
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://www.duifene.com/_CheckIn/PC/StudentNoCheckCount.aspx"
    }
    print(f"student_id={student_id}, class_id={Course.class_id}")
    try:
        print("Sending request...")
        _r = x.post(host + "/_CheckIn/MBCount.ashx", data=post_data, headers=headers, timeout=10)
        print(f"Response status: {_r.status_code}")
        if _r.status_code != 200:
            text_box.insert(tk.END, f"\n请求签到列表失败，HTTP {_r.status_code}")
            if Course.flag:
                root.after(1000, watching_sign)
            return
        data = _r.json()
        print(f"Data: {data}")
        if data.get("msg") != "1":
            text_box.insert(tk.END, f"\n接口返回错误：{data.get('msgbox', '未知错误')}")
            if Course.flag:
                root.after(1000, watching_sign)
            return

        rows = data.get("rows", [])
        if not rows:
            if Course.flag:
                root.after(1000, watching_sign)
            return

        for item in rows:
            check_in_id = item.get("ID")
            if not check_in_id:
                continue
            if check_in_id in Course.check_list:
                continue
            if item.get("CanApply") != "1":
                continue

            create_date = item.get("CreaterDate")
            if not create_date:
                continue

            try:
                create_time = datetime.strptime(create_date, "%Y/%m/%d %H:%M:%S")
                now = datetime.now()
                # 如果创建时间比当前时间晚太多（超过5分钟），可能是错误数据，忽略
                if (create_time - now).total_seconds() > 300:
                    text_box.insert(tk.END, f"\n活动 {check_in_id} 创建时间 {create_date} 过于未来，忽略")
                    continue
                
                elapsed = (now - create_time).total_seconds()
                if elapsed < 15:
                    remaining_wait = int(15 - elapsed)
                    text_box.insert(tk.END, f"\n活动 {check_in_id} 还需等待 {remaining_wait} 秒才开始签到")
                    continue
                # elapsed >= 15，可以签到
            except Exception as e:
                text_box.insert(tk.END, f"\n活动 {check_in_id} 创建时间解析失败: {e}")
                continue

            check_in_type = item.get("CheckInType")
            status = False
            text_box.insert(tk.END, f"\n\n{current_time} 发现签到活动 ID:{check_in_id}（创建于 {create_date}，已过去 {int(elapsed)} 秒），开始签到...")

            if check_in_type == '1':
                sign_code = item.get("CheckInCode")
                if sign_code and len(sign_code) == 4:
                    text_box.insert(tk.END, f"\t签到码：{sign_code}")
                    status = sign(sign_code)
                else:
                    text_box.insert(tk.END, "\t无效签到码")
            elif check_in_type == '2':
                text_box.insert(tk.END, "\t二维码签到模式")
                status = sign(check_in_id)
            elif check_in_type == '3':
                text_box.insert(tk.END, "\t定位签到模式，尝试获取经纬度...")
                try:
                    loc_r = x.get(url=host + f"/_CheckIn/MB/TeachCheckIn.aspx?classid={Course.class_id}&temps=0&checktype=1&isrefresh=0&timeinterval=0&roomid=0&match=", timeout=10)
                    if loc_r.status_code == 200:
                        soup = BeautifulSoup(loc_r.text, "lxml")
                        lng_elem = soup.find(id="HFRoomLongitude")
                        lat_elem = soup.find(id="HFRoomLatitude")
                        if lng_elem and lat_elem:
                            longitude = lng_elem.get("value")
                            latitude = lat_elem.get("value")
                            if longitude and latitude:
                                status = sign_location(longitude, latitude)
                            else:
                                text_box.insert(tk.END, "\t定位签到失败：未获取到经纬度")
                        else:
                            text_box.insert(tk.END, "\t定位签到失败：页面缺少经纬度信息")
                    else:
                        text_box.insert(tk.END, f"\t定位签到失败：HTTP {loc_r.status_code}")
                except Exception as e:
                    text_box.insert(tk.END, f"\t定位签到异常：{str(e)}")
            else:
                text_box.insert(tk.END, f"\t未知签到类型 {check_in_type}，暂不支持")

            if status:
                Course.check_list.append(check_in_id)
                text_box.insert(tk.END, "\t签到成功记录已保存")
            else:
                text_box.insert(tk.END, "\t签到失败，请查看上方信息")

    except requests.exceptions.RequestException as e:
        text_box.insert(tk.END, f"\n网络请求异常：{str(e)}")
    except Exception as e:
        text_box.insert(tk.END, f"\n处理签到列表时出错：{str(e)}")

    # 继续循环监控
    if Course.flag:
        root.after(1000, watching_sign)


def go_sign():
    if combo.get() is None or combo.get() == '':
        messagebox.showerror("错误提示", "请先登录")
        return
    if Course.id == '0' or Course.class_id == '0':
        messagebox.showerror("错误提示", "课程未正确加载，请重新登录")
        return

    text_box.delete("1.0", "end")
    CourseName = combo_var.get()
    text_box.insert(tk.END, f"正在监听【{CourseName}】的签到活动\n\n")
    watching_sign()


def get_class_list():
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/_UserCenter/PC/CenterStudent.aspx"
    }
    params = "action=getstudentcourse&classtypeid=2"
    _r = x.post(url=host + "/_UserCenter/CourseInfo.ashx", data=params, headers=headers)
    if _r.status_code == 200:
        _json = _r.json()
        if _json is not None:
            try:
                msg = _json["msgbox"]
                messagebox.showerror("", f"{msg} 请重新登录。")
                x.cookies.clear()
            except Exception:
                messagebox.showinfo("提示", "登录成功")
                class_name_list = []
                for i in _json:
                    class_name_list.append(i["CourseName"])
                combo['values'] = tuple(class_name_list)
                combo.set(class_name_list[0])
                Course.id = _json[0]['CourseID']
                Course.class_id = _json[0]["TClassID"]
                Course.class_list = _json


def is_login():
    headers = {
        "Referer": "https://www.duifene.com/_UserCenter/PC/CenterStudent.aspx",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    }
    _r = x.get(host + "/AppCode/LoginInfo.ashx", data="Action=checklogin", headers=headers)
    if _r.status_code == 200:
        if _r.json()["msg"] == "1":
            return True
        else:
            messagebox.showwarning("登录状态失效", "请重新登录账号")
            x.cookies.clear()
            Course.flag = False
            return False


def init():
    try:
        if not os.path.exists(filename):
            config['INFO'] = {
                'cookie': '1=1'
            }
            with open(filename, 'w') as configfile:
                config.write(configfile)
            x.get(host)
        else:
            try:
                config.read(filename)
                cookie = config.get('INFO', 'cookie')
                cookies = {}
                for pair in cookie.split('; '):
                    key, value = pair.split('=')
                    cookies[key] = value
                x.cookies.update(cookies)
                get_class_list()
            except Exception:
                pass
    except (requests.ConnectionError, requests.Timeout):
        messagebox.showwarning("网络状态", "未检测到互联网连接，请检查你的网络设置。")
        root.destroy()


if __name__ == '__main__':
    host = "https://www.duifene.com"
    UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) ' \
         'Mobile/15E148 MicroMessenger/8.0.40(0x1800282a) NetType/WIFI Language/zh_CN '
    urllib3.disable_warnings()
    x = requests.Session()
    x.headers['User-Agent'] = UA
    x.verify = False
    filename = 'duifenyi.ini'
    config = configparser.ConfigParser()

    root = tk.Tk()
    root.title("2024.9.18")
    root.resizable(False, False)

    tab_control = ttk.Notebook(root)
    tab1 = ttk.Frame(tab_control)
    tab2 = ttk.Frame(tab_control)
    tab_control.add(tab1, text="微信链接登录")
    tab_control.add(tab2, text="账号密码登录")
    tab_control.bind("<<NotebookTabChanged>>", select_tab)
    tab_control.pack(fill=tk.BOTH, side=tk.LEFT)

    tab_frame1 = tk.Frame(tab_control)
    tab_frame1.pack(side=tk.LEFT, fill=tk.BOTH, pady=(40, 0))
    tk.Label(tab_frame1, text="支持二维码和签到码\n查看右侧说明进行登录", font=('宋体', 10)).pack(pady=5)
    tk.Label(tab_frame1, text="登录链接", font=('宋体', 10)).pack(pady=5)
    link_entry = tk.Entry(tab_frame1, font=('宋体', 12))
    link_entry.pack(pady=5, padx=10)
    tk.Button(tab_frame1, text="登录", command=login_link, font=('宋体', 14)).pack(pady=5)

    tab_frame2 = tk.Frame(tab_control)
    tk.Label(tab_frame2, text="不支持二维码签到", font=('宋体', 10)).pack(padx=5)
    tk.Label(tab_frame2, text="账号", font=('宋体', 14)).pack(padx=10)
    username = tk.Entry(tab_frame2, font=('宋体', 12))
    username.pack(padx=10)
    tk.Label(tab_frame2, text="密码", font=('宋体', 14)).pack(padx=10)
    password = tk.Entry(tab_frame2, show="*", font=('宋体', 12))
    password.pack(padx=10)
    tk.Label(tab_frame2, text="剩余倒计时X秒后签到", font=('宋体', 10)).pack(pady=5)
    seconds_entry = tk.Entry(tab_frame2, font=('宋体', 12), width=5)
    seconds_entry.insert(0, "10")
    seconds_entry.pack(pady=5)
    tk.Button(tab_frame2, text="登录", command=login, font=('宋体', 14)).pack(pady=5)

    frame_mid = tk.Frame(root)
    frame_mid.pack(side=tk.TOP)
    tk.Label(frame_mid, text="选择课程").pack(side=tk.TOP, fill=tk.BOTH, pady=(10, 0))
    combo_var = tk.StringVar()
    combo = ttk.Combobox(frame_mid, textvariable=combo_var, state="readonly")
    combo.bind("<<ComboboxSelected>>", on_combo_change)
    combo.pack(side=tk.LEFT)
    btn = tk.Button(frame_mid, text="开始监听签到", command=go_sign)
    btn.pack(side=tk.RIGHT, padx=10, pady=10)

    frame_right = tk.Frame(root)
    frame_right.pack(side=tk.RIGHT)
    text_box = tk.Text(frame_right, width=90, height=20, font=('宋体', 9))
    text_box.pack(pady=(0, 10), padx=(0, 10))

    init()
    root.mainloop()