import json
import time
import ctypes
from selenium import webdriver
from threading import Thread

import ttkbootstrap as ttk
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager


class ZhiHuiShuTool:
    def __init__(self):
        self.driver = None
        self.root = None
        self.ipturl = None
        self.iptuser = None
        self.iptpwd = None
        self.textbox = None

    def quit_app(self):
        try:
            if self.driver:
                self.driver.quit()
            self.root.destroy()
        except Exception as e:
            self.root.destroy()

    def setup_driver(self):
        self.log_message("驱动检查中，第一次运行可能会较慢")
        service = Service(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service)
        self.log_message("完成，准备启动")
        self.driver = driver

    def log_message(self, message, end="\n"):
        self.textbox.insert("1.0", message + end)

    def login(self, user, pwd, url):
        self.driver.get(url)
        try:
            username_field = self.driver.find_element(By.ID, "lUsername")
            password_field = self.driver.find_element(By.ID, "lPassword")
            username_field.send_keys(user)
            password_field.send_keys(pwd)
        except Exception as e:
            pass

        if not user or not pwd:
            self.log_message("账号密码未填写，请自行登录")
        else:
            try:
                submit_button = self.driver.find_element(By.CLASS_NAME, "wall-sub-btn")
                submit_button.click()
            except Exception as e:
                self.log_message("登录按钮点击失败")

    def wait_for_login(self):
        success = False
        for i in range(1, 100):
            self.log_message(f'等待登录，第{i}次')
            time.sleep(3)
            videos = self.driver.find_elements(By.CLASS_NAME, "clearfix.video")
            if videos:
                self.log_message("登录成功")
                success = True
                break

        if not success:
            self.log_message("code = 0；错误：登录超时，请重试")
            self.driver.quit()

    def play_videos(self):
        self.log_message('开始执行')
        self.driver.execute_script(
            'javascript:v = function(){a = document.getElementById("playTopic-dialog");a.remove();b = document.getElementsByClassName("v-modal")[0];b.remove()};'
            'javascript:setInterval(v,1000);'
            'setInterval(function(){var pruse=document.getElementsByTagName("video");for (var i = 0; i < pruse.length; i ++){pruse[i].play()}},1000);'
        )

        video_elements = self.driver.find_elements(By.CLASS_NAME, "clearfix.video")
        self.log_message(f"加载课程成功，共计 {len(video_elements)} 集")

        while True:
            time.sleep(5)
            try:
                total_time = self.driver.execute_script('return document.getElementsByTagName("video")[0].duration;')
                current_time = self.driver.execute_script(
                    'return document.getElementsByTagName("video")[0].currentTime;')
                progress = round(current_time / total_time, 2) if total_time else 0
                self.log_message(f"总时间：{round(total_time)}秒 时间：{round(current_time)}秒 进度：{progress}")

                if progress > 0.9:
                    self.log_message("进度大于0.9，跳转下一集")
                    try:
                        video_elements = self.driver.find_elements(By.CLASS_NAME, "clearfix.video")
                        current_video = next((i for i in video_elements if "current_play" in i.get_attribute("class")),
                                             None)
                        if current_video and video_elements.index(current_video) < len(video_elements) - 1:
                            next_video = video_elements[video_elements.index(current_video) + 1]
                            next_video.click()
                            title = next_video.find_element(By.TAG_NAME, "span").text
                            self.log_message(f"播放 {title}")
                        else:
                            self.log_message("已经是最后一集，播放结束")
                            self.quit_app()
                            break
                    except Exception as e:
                        self.log_message("获取下一集失败")
            except Exception as e:
                continue

    def start_browser(self):
        user = self.iptuser.get()
        pwd = self.iptpwd.get()
        url = self.ipturl.get() or "https://passport.zhihuishu.com/login"

        with open("configZD.json", "w") as f:
            f.write(json.dumps({"user": user, "pwd": pwd, "url": url}))

        thread = Thread(target=self.main, args=(user, pwd, url), daemon=True)
        thread.start()

    def main(self, user, pwd, url):
        self.setup_driver()
        self.login(user, pwd, url)
        self.wait_for_login()
        self.play_videos()

    def setup_ui(self):
        self.root = ttk.Window(title="智慧树", themename="litera", size=(450, 590), resizable=None)
        ttk.Label(self.root, text="知道刷课工具").place(x=70, y=0, width=300, height=30)
        ttk.Label(self.root, text="课程链接").place(x=20, y=40, width=80, height=30)

        self.ipturl = ttk.Entry(self.root)
        self.ipturl.place(x=111, y=35, width=310, height=35)

        self.iptuser = ttk.Entry(self.root)
        self.iptuser.place(x=80, y=80, width=340, height=30)

        label = ttk.Label(self.root, text="账号", anchor="center")
        label.place(x=20, y=80, width=50, height=30)

        label = ttk.Label(self.root, text="密码", anchor="center")
        label.place(x=20, y=120, width=50, height=30)

        self.iptpwd = ttk.Entry(self.root)
        self.iptpwd.place(x=80, y=120, width=340, height=30)

        self.textbox = ttk.Text(self.root)
        self.textbox.place(x=30, y=180, width=390, height=325)

        self.textbox.insert("1.0", "课程链接不填默认到智慧树官网，请自行打开课程页\n")
        self.textbox.insert("1.0", "账号密码不写可在浏览器中自己登录\n")
        self.textbox.insert("1.0", "注意不要全屏播放视频，否则会无法自动换集\n")

        ttk.Button(self.root, text="开始", command=self.start_browser).place(x=35, y=520, width=100, height=40)
        ttk.Button(self.root, text="关闭", command=self.quit_app).place(x=315, y=520, width=92, height=40)

    def run(self):
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100  # 从 100% 换算为 1.0, 125% 为 1.25
        self.root.tk.call('tk', 'scaling', scale_factor)  # 设置窗口的缩放
        try:
            with open("configZD.json", "r") as f:
                config = json.load(f)
                self.ipturl.insert("end", config.get("url", ""))
                self.iptuser.insert("end", config.get("user", ""))
                self.iptpwd.insert("end", config.get("pwd", ""))
        except Exception as e:
            self.log_message("读取配置文件失败")
        self.root.mainloop()


if __name__ == '__main__':
    app = ZhiHuiShuTool()
    app.setup_ui()
    app.run()
