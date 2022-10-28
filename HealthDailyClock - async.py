import os
import csv
import time
import shutil
import random
import zipfile
import asyncio
import selenium
import requests
from PIL import Image
from chaojiying import Chaojiying_Client
from selenium.webdriver import ChromeOptions


# 检查谷歌版本并矫正
def ChromeExam():
    if '92.0.4515.107' not in os.listdir(r'C:\Program Files\Google\Chrome\Application'):
        os.startfile(r'C:\Program Files\Google\Chrome\Application')
        print('请删除，回车继续')
        input()
        chrome_zip = r'D:\anyt\stud\个人项目\[不可用]成绩自动查询助手\Chrome-bin\Chrome-bin.zip'  # 压缩包~
        chrome_dir = r'C:\Program Files\Google\Chrome\Application'  # 目标文件夹
        for file in os.listdir(chrome_dir):  # 把目标文件夹清空
            current_path = os.path.join(chrome_dir, file)  # 权限问题：把要操作的文件夹安全改一改就可以了
            if os.path.isdir(current_path):
                shutil.rmtree(current_path)  # 删文件夹及所有子文件
            else:
                os.remove(current_path)  # 删除文件
        shutil.copyfile(chrome_zip, os.path.join(chrome_dir, 'Chrome-bin.zip'))  # 复制文件夹
        zip_file = zipfile.ZipFile(os.path.join(chrome_dir, 'Chrome-bin.zip'))  # 扫描压缩包（参数为压缩包路径）
        for name in zip_file.namelist():
            zip_file.extract(name, chrome_dir)  # 一一解压到chrome_dir中
        zip_file.close()
        os.remove(os.path.join(chrome_dir, 'Chrome-bin.zip'))  # 把压缩包也删除吧
        print('--------------矫正了Chrome版本-----------------')
    else:
        print('---------------Chrome版本正常-----------------')


# 判断元素是否存在
def isElement(bro, xpath):
    try:
        bro.find_element_by_xpath(xpath)
        return True
    except selenium.common.exceptions.NoSuchElementException:
        return False


# 读取数据库
def read_database():
    database = dict()
    with open(r'D:\anyt\stud\个人项目\健康日报打卡\database.csv') as f:
        f_csv = csv.reader(f)
        headers = next(f_csv)
        for row in f_csv:
            database[row[0]] = {'username': row[1], 'password': row[2], 'token': row[3], 'flag': False}
    return database


class HealthClock:
    def __init__(self):
        ChromeExam()  # 谷歌版本矫正
        self.database = read_database()
        self.clock_text = ""
        self.error_text = ""
        self.cnt = 0

    def push_message(self, user, title, content):
        if self.database[user]['token'] != '':
            initial_data = {'token': self.database[user]['token'],
                            'title': title,
                            'content': content}
            requests.post(url="http://www.pushplus.plus/send", data=initial_data)

    # ----------------------------------------------核心代码--------------------------------------------------
    async def ClockSingle(self, user):
        # 未打卡判断
        if not self.database[user]['flag']:
            print('{} 打卡中..'.format(user))

            # webdriver配置
            options = ChromeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
            bro = selenium.webdriver.Chrome(executable_path=r'D:\anyt\stud\个人项目\健康日报打卡\chromedriver.exe',
                                            options=options)

            # 让我访问
            bro.get('http://i.nuist.edu.cn/qljfwapp/sys/lwNuistHealthInfoDailyClock/index.do#/healthClock')

            # 直至验证码输对
            await asyncio.sleep(1)

            # 验证码定位截取保存
            location = bro.find_element_by_id("captchaImg").location
            size = bro.find_element_by_id('captchaImg').size
            left = location['x'] + 180
            top = location['y'] + 70
            right = left + size['width'] + 50
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

            if 'login' in bro.current_url:
                bro.quit()
                print('{} 验证码错误，稍后重试'.format(user))
                self.error_text += '\n第{}轮，{} 打卡报错：\n验证码错误\n'.format(self.cnt, user)
                return

            await asyncio.sleep(8)

            # ----------------------------成功进入健康日报页面--------------------------------
            try:
                # 点击新增按钮
                bro.find_element_by_xpath('/html/body/main/article/section/div[2]/div[1]').click()
                time.sleep(3)

                # 若已打卡
                if isElement(bro, '/html/body/div[11]/div[1]/div[1]/div[1]'):
                    print('{} 今日已打卡'.format(user))
                    self.clock_text += '{} 今日已打卡\n'.format(user)
                    self.database[user]['flag'] = True
                    bro.quit()
                    return

                # 若未打卡
                else:
                    # 填写温度
                    bro.find_element_by_xpath(
                        '/html/body/div[11]/div/div[1]/section/div[2]/div/div[3]/div[2]/div[18]/div/div/input').send_keys(
                        str(round(random.uniform(36.0, 36.5), 1)))  # 填写温度

                    time.sleep(1)

                    # 点击保存
                    bro.find_element_by_xpath('/html/body/div[11]/div/div[2]/footer/div').click()

                    time.sleep(2)

                    # 点击确认
                    bro.find_element_by_xpath('/html/body/div[30]/div[1]/div[1]/div[2]/div[2]/a[1]').click()

                    # 打卡成功啦
                    bro.quit()

                    print('\r{} 打卡成功'.format(user))
                    self.clock_text += '{} 打卡成功\n'.format(user)
                    self.database[user]['flag'] = True
                    self.push_message(user=user,
                                      title='健康日报打卡成功',
                                      content='{} 健康日报打卡成功'.format(user))

            except Exception as error:
                bro.quit()
                print('{} 打卡报错：\n{}'.format(user, error))
                self.error_text += '\n第{}轮，{} 打卡报错：\n{}\n'.format(self.cnt, user, error)

    # -------------------------------------------------------------------------------------------------------

    def ClockTogether(self):
        tasks = []
        for user in self.database:
            if not self.database[user]['flag']:
                tasks.append(asyncio.ensure_future(self.ClockSingle(user)))
        loop = asyncio.get_event_loop()
        loop.run_until_complete(asyncio.wait(tasks))

    def ClockExamine(self):
        for user in self.database:
            if not self.database[user]['flag']:
                return True  # 存在没打卡成功的
        return False  # 不存在没打卡成功的

    def AutoClock(self):
        start = time.time()
        while self.ClockExamine() and self.cnt <= 5:
            self.cnt += 1
            print("-------第{}遍打卡-------".format(self.cnt))
            self.ClockTogether()
        end = time.time()

        if self.cnt <= 5:
            print('-------------------健康日报打卡完成!-----------------------')
            print('打卡情况将通过微信推送，', end='')
            print('本次耗时：{}s'.format(round(end - start)))
            self.clock_text += '本次耗时：{}s'.format(round(end - start))
            self.push_message(user='刘安邦',
                              title='今日打卡情况汇报',
                              content=self.clock_text)
        else:
            print('~~~~~~~~~~~~~~~~~~~~打卡程序出现问题~~~~~~~~~~~~~~~~~~~~~~~')
            self.push_message(user='刘安邦',
                              title='打卡程序出现问题',
                              content=self.error_text)


if __name__ == '__main__':
    hc = HealthClock()
    hc.AutoClock()
