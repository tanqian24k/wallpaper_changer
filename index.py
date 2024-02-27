import os
import sys
import time
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QSystemTrayIcon, QMenu, QAction, qApp, QLabel, QFileDialog
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import QTimer, Qt
import random
import ctypes
import requests
import configparser

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
    "Wallpaper Changer")


class mainApp:
    def __init__(self):
        # 获取exe所在的目录
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(
            os.path.abspath(__file__)))
        resources_path = os.path.join(base_path, 'resources')

        print(resources_path)
        self.app = QApplication(sys.argv)

        self.app.setApplicationName('Wallpaper Changer')
        self.app.setApplicationDisplayName('Wallpaper Changer')

        self.checkedIcon = QIcon(os.path.join(resources_path, 'checked.png'))
        self.refreshIcon = QIcon(os.path.join(resources_path, 'reload.png'))
        self.icon = QIcon(os.path.join(resources_path, 'icon.png'))

        self.app.setWindowIcon(self.icon)
        self.app.setQuitOnLastWindowClosed(False)

        self.payment = QMainWindow()
        self.payment.setWindowTitle('赞助作者')
        self.payment.setWindowFlags(Qt.WindowCloseButtonHint)
        # 设置任务栏图标

        self.payment.setWindowIcon(self.icon)
        self.payment.resize(450, 450)

        # 创建一个 QLabel 部件
        label = QLabel(self.payment)
        # 将 QPixmap 设置为 QLabel 的内容
        # 设置QPixmap的尺寸
        label.setScaledContents(True)
        # 设置QPixmap的尺寸
        label.setFixedSize(450, 450)
        label.setPixmap(QPixmap(os.path.join(resources_path, 'payment.jpg')))
        # 将 QLabel 设置为 QMainWindow 的中心部件
        self.payment.setCentralWidget(label)

        self.tray_icon = QSystemTrayIcon(self.icon, self.app)
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        self.sourceUrl = "https://source.unsplash.com/random/"
        self.styleTags = self.config.get('tags', 'style').split(',')
        # 移除空格
        self.styleTags = [tag.strip() for tag in self.styleTags]
        self.sceneTags = self.config.get('tags', 'scene').split(',')
        self.sceneTags = [tag.strip() for tag in self.sceneTags]
        self.width = self.app.desktop().width()
        self.height = self.app.desktop().height()
        self.menu = QMenu()
        self.timer = QTimer()
        self.timer.timeout.connect(lambda: self.refresh_wallpaper)
        self.gen_menu()
        self.tray_icon.setContextMenu(self.menu)
        self.tray_icon.show()
        self.init_timer()

    def gen_menu(self):
        self.menu.clear()
        self.menu.addAction(QAction(self.refreshIcon, '刷新',
                            self.app, triggered=self.refresh_wallpaper))
        self.menu.addAction(
            QAction('保存当前壁纸', self.app, triggered=self.save_wallpaper))
        self.menu.addSeparator()
        isenabled = self.config.get('timer', 'enable')
        config_timeval = self.config.get('timer', 'interval')
        # define time interval
        # read from config file
        time_interval = self.config.get('timer', 'options')
        if time_interval != '':
            # convert to int
            time_interval = [int(i) for i in time_interval.split(',')]
        else:
            time_interval = [15, 30, 60, 120, 240, 720, 1440]
        for i in time_interval:
            temp_action = QAction(
                str(i)+'分钟' if i < 60 else str(i//60)+'小时', self.app,
                triggered=lambda x, y=i: self.set_timer(y)
            )
            if (isenabled == '1' and config_timeval == str(i)):
                temp_action.setIcon(self.checkedIcon)
            self.menu.addAction(temp_action)

        stop_action = QAction('停止自动刷新', self.app,
                              triggered=lambda: self.set_timer(0))
        if isenabled == '0':
            stop_action.setIcon(self.checkedIcon)
        self.menu.addAction(stop_action)
        self.menu.addSeparator()
        self.menu.addAction(
            QAction('赞助作者', self.app, triggered=lambda: self.payment.show()))
        self.menu.addAction(QAction('退出', self.app, triggered=self.close))

    def refresh_wallpaper(self):
        tag = random.choice(self.styleTags)+','+random.choice(self.sceneTags)
        print(tag)
        url = self.sourceUrl + str(self.width) + \
            'x' + str(self.height) + '/?' + tag
        self.set_wallpaper_from_url(url)

    def save_wallpaper(self):
        # 获取当前目录
        current_path = os.getcwd()
        # 获取图片路径
        image_path = os.path.join(current_path, 'wallpaper.jpg')
        # 读取配置文件中目标路径
        save_path = self.config.get('save', 'path')
        # 如果配置文件中没有目标路径，则弹出对话框选择保存路径
        if save_path == '':
            # save_path = filedialog.askdirectory()
            save_path = QFileDialog.getExistingDirectory(None, "选取文件夹")
            # 将选择的路径保存到配置文件中
            self.config.set('save', 'path', save_path)
            with open('config.ini', 'w') as configfile:
                self.config.write(configfile)
        if save_path == '':
            return
        # 获取当前时间
        current_time = time.strftime(
            '%Y%m%d%H%M%S', time.localtime(time.time()))
        # 复制图片到指定目录并修改名称
        shutil.copy(image_path, os.path.join(save_path, current_time + '.jpg'))
        # 资源管理器打开到当前目录
        os.startfile(save_path)

    def set_wallpaper_from_url(self, url):
        response = requests.get(url)
        image_path = 'wallpaper.jpg'
        if response.status_code == 200:
            with open(image_path, 'wb') as file:
                file.write(response.content)
            print("图片保存成功！")
        else:
            print("图片下载失败！")
        image_path = os.path.join(os.getcwd(), image_path)
        # 使用ctypes库设置桌面壁纸
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 3)

    def set_timer(self, interval):
        if (interval == 0):
            self.config.set('timer', 'enable', '0')
        else:
            self.config.set('timer', 'enable', '1')
            self.config.set('timer', 'interval', str(interval))
        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)
        self.gen_menu()
        self.init_timer()

    def init_timer(self):
        if self.timer != None and self.timer.isActive():
            self.timer.stop()
        # 读取配置
        timerEnable = self.config.get('timer', 'enable')
        if timerEnable == "0":
            return
        self.refresh_wallpaper()
        interval = int(self.config.get('timer', 'interval'))*60000
        self.timer = QTimer()
        self.timer.timeout.connect(self.refresh_wallpaper)
        self.timer.setInterval(interval)
        self.timer.start()

    def run(self):
        self.app.exec_()

    def close(self):
        self.tray_icon.hide()
        self.app.quit()


if __name__ == '__main__':
    app = mainApp()
    app.run()
