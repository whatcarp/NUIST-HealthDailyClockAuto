import os
import csv
import time
import random
import selenium
import requests
import schedule
import datetime
from PIL import Image
from chaojiying import Chaojiying_Client
from selenium.webdriver import ChromeOptions


class HealthClock:

    def __init__(self):
        self.OK = True
        self.logs = ""
        self.database = self.read_database()

    """-----------------------------------------------单人打卡--------------------------------------------------------"""

    def clock_single(self, user):

        print('\r{} 打卡中..'.format(user), end='')

        # webdriver配置
        options = ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
        bro = selenium.webdriver.Chrome(executable_path='./chromedriver.exe', options=options)

        # 让我访问
        bro.get('http://i.nuist.edu.cn/qljfwapp/sys/lwNuistHealthInfoDailyClock/index.do#/healthClock')

        # ---------------------------------------登录环节---------------------------------------------
        cnt = 0
        while 'login' in bro.current_url:
            time.sleep(3)
            # 验证码定位截取保存
            bro.find_element_by_id('username').clear()
            bro.find_element_by_id('password').clear()
            bro.find_element_by_id('username').send_keys(self.database[user]['username'])
            bro.find_element_by_id('password').send_keys(self.database[user]['password'])
            bro.find_element_by_id('login_submit').click()

            location = bro.find_element_by_id("captchaImg").location
            size = bro.find_element_by_id('captchaImg').size
            left = location['x'] + 205
            top = location['y'] + 70
            right = left + size['width'] + 25
            bottom = top + size['height'] + 10
            captcha_img_path = "./CaptchaImg.png"
            bro.get_screenshot_as_file(captcha_img_path)  # 截屏
            img = Image.open(captcha_img_path).crop((left, top, right, bottom))
            img.save(captcha_img_path)

            # 验证码识别
            captcha_recognize = Chaojiying_Client('shilizi', 'lab20010119', '921485')
            captcha_img = open(captcha_img_path, 'rb').read()
            captcha = captcha_recognize.PostPic(captcha_img, 1902)['pic_str']
            os.remove('./CaptchaImg.png')

            # 填写信息
            bro.find_element_by_id('username').clear()
            bro.find_element_by_id('password').clear()
            bro.find_element_by_id('captcha').clear()
            bro.find_element_by_id('username').send_keys(self.database[user]['username'])
            bro.find_element_by_id('password').send_keys(self.database[user]['password'])
            bro.find_element_by_id('captcha').send_keys(captcha)
            bro.find_element_by_id('login_submit').click()

            if 'login' in bro.current_url:  # 验证码输入错误
                cnt += 1
                if cnt > 3:  # 验证码输错3次
                    print('{} 验证码验证失败3次，暂时跳过\n'.format(user))
                    bro.quit()
                    return
        # -------------------------------------------------------------------------------------------

        time.sleep(10)  # 小憩一下

        # -----------------------------------成功进入健康日报页面----------------------------------------
        try:
            # 点击新增按钮
            bro.find_element_by_xpath('/html/body/main/article/section/div[2]/div[1]').click()
            time.sleep(5)

            # 已打卡
            if self.isElement(bro, '/html/body/div[11]/div[1]/div[1]/div[1]'):
                print('\r{} 今日已打卡'.format(user))
                self.logs += '{} 今日已打卡\n'.format(user)
                self.database[user]['flag'] = True

                bro.quit()
                return

            # 未打卡
            else:
                # 填写温度
                bro.find_element_by_xpath(
                    '/html/body/div[11]/div/div[1]/section/div[2]/div/div[3]/div[2]/div[18]/div/div/input').send_keys(
                    str(round(random.uniform(36.0, 36.5), 1)))  # 填写温度
                time.sleep(2)
                # 点击保存
                bro.find_element_by_xpath('/html/body/div[11]/div/div[2]/footer/div').click()
                time.sleep(1)
                # 点击确认
                bro.find_element_by_xpath('/html/body/div[30]/div[1]/div[1]/div[2]/div[2]/a[1]').click()
                time.sleep(1)

                # 打卡成功啦
                print('\r{} 打卡成功'.format(user))
                self.logs += '{} 打卡成功\n'.format(user)
                self.database[user]['flag'] = True

                # 给用户发送打卡成功消息
                self.push_message(user=user,
                                  title='健康日报打卡成功',
                                  content='{}\n{} 健康日报打卡成功'.format(datetime.datetime.now(), user))

                bro.quit()
                return

        except Exception as error:
            print('{} 打卡部分程序报错，暂时跳过\n'.format(user), error)
            bro.quit()
            return

    """--------------------------------------------------------------------------------------------------------------"""

    # 按时打卡
    def clock_on_time(self):
        print("\n\n{} 健康日报自动打卡启动\n\n".format(datetime.datetime.now()))
        schedule.every().day.at("08:00").do(self.clock_together)
        while True:
            schedule.run_pending()
            time.sleep(1)

    # 统一打卡
    def clock_together(self):
        self.__init__()  # 初始化一下当前状态
        print('\n\n\n------------{}------------\n'.format(datetime.datetime.now().date()))
        cnt = 0
        while self.clock_examine():
            cnt += 1
            if cnt > 10:
                self.OK = False
                break

            print('-----------第{}次打卡----------'.format(cnt))
            for user in self.database:
                if not self.database[user]['flag']:
                    self.clock_single(user)

        if self.OK:
            print("{} 全员打卡完成".format(datetime.datetime.now()))
            self.push_message(user='刘安邦', title='今日打卡情况汇报', content=self.logs)

        else:
            print('打卡错误次数过多，疑似出现问题，将发送消息至管理员，并终止今日打卡')
            self.push_message(user='刘安邦',
                              title='今日打卡出现问题',
                              content='打卡错误次数过多，疑似出现问题，已终止今日打卡\n' + self.logs)

    # 检查打卡情况
    def clock_examine(self):
        for user in self.database:
            if not self.database[user]['flag']:
                return True  # 存在没打卡成功的
        return False  # 不存在没打卡成功的

    # 消息推送封装
    def push_message(self, user, title, content):
        if self.database[user]['token'] != '':
            initial_data = {'token': self.database[user]['token'],
                            'title': title,
                            'content': content}
            requests.post(url="http://www.pushplus.plus/send", data=initial_data)

    # 读取csv格式数据库
    @staticmethod
    def read_database():
        database = dict()
        with open('./database.csv') as f:
            f_csv = csv.reader(f)
            headers = next(f_csv)
            for row in f_csv:
                database[row[0]] = {'username': row[1], 'password': row[2], 'token': row[3], 'flag': False}
        return database

    # 判断元素是否存在
    @staticmethod
    def isElement(bro, xpath):
        try:
            bro.find_element_by_xpath(xpath)
            return True
        except selenium.common.exceptions.NoSuchElementException:
            return False


if __name__ == '__main__':
    hc = HealthClock()
    hc.clock_on_time()
