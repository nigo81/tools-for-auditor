#!/usr/bin/python
# -*- coding: utf-8 -*-
# Author: wwcheng

import csv
import datetime
import logging
import logging.handlers
import os
import random
import re
import time
from urllib.request import urlretrieve

import pandas as pd
import prettytable as pt
import requests
from PIL import Image
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class Express_Tracking(object):

    def __init__(self,count=0):
        self.query_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d %H-%M-%S')
        self.nowpath = os.path.abspath(os.curdir)
        self.sf_Express_url = r'http://www.sf-express.com/cn/sc/dynamic_function/waybill/#search/bill-number/'
        self.kuaidi100_url = r'https://www.kuaidi100.com/?from=openv'
        self.html_size=0.7
        self.mypath=[]
        self.sf_tab=''
        self.kd100_tab=''
        self.count=count
        self.route=''
        self.dzcg=''
        self.shot_name=''
        self.__WelCome()
        self.__ceart_enviroment()
        self.__logging()
        
        if 'PROGRAMFILES(X86)' in os.environ:
            self.executable_path = self.nowpath+r'\Driver\geckodriver-win64.exe'
        else:
            self.executable_path = self.nowpath+r'\Driver\geckodriver-win32.exe'
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.add_argument('--headless')
        self.browser = webdriver.Firefox(
            executable_path=self.executable_path,
            firefox_options=firefox_options,
            service_log_path=self.mypath[3]+r'\geckodriver.log'
        )
        
        self.__open_browser()
        self.__scan_login()
        self.user_agent_list = [
                # Opera
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36 OPR/26.0.1656.60",
                "Opera/8.0 (Windows NT 5.1; U; en)",
                "Mozilla/5.0 (Windows NT 5.1; U; en; rv:1.8.1) Gecko/20061208 Firefox/2.0.0 Opera 9.50",
                "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; en) Opera 9.50",
                # Firefox
                "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0",
                "Mozilla/5.0 (X11; U; Linux x86_64; zh-CN; rv:1.9.2.10) Gecko/20100922 Ubuntu/10.10 (maverick) Firefox/3.6.10",
                # Safari
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/534.57.2 (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2",
                # chrome
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.71 Safari/537.36",
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11",
                "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.16 (KHTML, like Gecko) Chrome/10.0.648.133 Safari/534.16",
                # 360
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.101 Safari/537.36",
                "Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko",
                # 淘宝浏览器
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.11 TaoBrowser/2.0 Safari/536.11",
                # 猎豹浏览器
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/21.0.1180.71 Safari/537.1 LBBROWSER",
                "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; LBBROWSER)",
                "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; QQDownload 732; .NET4.0C; .NET4.0E; LBBROWSER)",
                # QQ浏览器
                "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; QQBrowser/7.0.3698.400)",
                "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; QQDownload 732; .NET4.0C; .NET4.0E)",
                # sogou浏览器
                "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.84 Safari/535.11 SE 2.X MetaSr 1.0",
                "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SV1; QQDownload 732; .NET4.0C; .NET4.0E; SE 2.X MetaSr 1.0)",
                # maxthon浏览器
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Maxthon/4.4.3.4000 Chrome/30.0.1599.101 Safari/537.36",
                # UC浏览器
                "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.2125.122 UBrowser/4.0.3214.0 Safari/537.36",
            ]

    def __logging(self):
        # 内部函数；用于记录日志
        logging.basicConfig(filename=self.mypath[3]+r"\Express_Tracking.log",
                            filemode="w",
                            format="%(asctime)s %(name)s:%(levelname)s:%(message)s",
                            datefmt="%d-%M-%Y %H:%M:%S",
                            level=logging.INFO
                            )
        self.logger_error = logging.getLogger("Express_Tracking")
        self.logger = logging.getLogger("15Scends")
        handler = logging.StreamHandler()
        handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            "%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.info('请在60秒内用手机完成微信扫码，超时后会视同不扫码继续运行')
        self.logger.info('可以不使用对应手机号的微信扫码，但是将无法获得电子存根信息')
        self.logger.info('成功扫码之后可以将二维码图片关闭')
        self.logger.info('正在打开微信扫码界面,请耐心等待30秒左右···\n')

    def debug(func):
        def wrapper(self,*args, **kwargs):
            try:
                return func(self,*args, **kwargs)
            except Exception:
                self.logger_error.exception("Recording Errors")
        return wrapper

    def __wait(self):
        # 内部函数；随机等待
        sleep_time = round(random.uniform(0.5, 2), 2)
        time.sleep(sleep_time)

    def __is_visible(self, locator, timeout=60):
        # 内部函数；判断网页元素是否出现
        try:
            WebDriverWait(self.browser, timeout).until(
                EC.visibility_of_element_located((By.XPATH, locator)))
            return True
        except TimeoutException:
            return False

    def __is_not_visible(self, locator, timeout=60):
        # 内部函数；判断网页元素是否消失
        try:
            WebDriverWait(self.browser, timeout).until_not(
                EC.visibility_of_element_located((By.XPATH, locator)))
            return True
        except TimeoutException:
            return False

    def __WelCome(self):
        # 打印程序描述
        tb = pt.PrettyTable()
        tb.field_names = ['名称', '快递批量查询截图工具2.0']
        tb.add_row(['作者', '王文铖'])
        tb.add_row(['微信公众号', '15Seconds'])
        tb.add_row(['', ''])
        tb.add_row(['', '关注微信公众号'])
        tb.add_row(['使用方法', '可查看工具详细使用方法'])
        tb.add_row(['', '还能获取更多实用工具'])
        tb.add_row(['', ''])
        tb.add_row(['GitHub项目地址', 'https://github.com/nigo81/tools-for-auditor'])
        tb.add_row(['', ''])
        tb.add_row(['更新时间', '2019-12-02'])
        print(tb)
        print('\n')

    def __ceart_enviroment(self):
        # 创建本次输出的文件环境
        result_path = self.nowpath+'\\Output\\' + self.query_time
        img_path = result_path+r'\img'
        xls_path = result_path+r'\xls'
        log_path = result_path+r'\log'
        temp_path = result_path+r'\temp'
        self.mypath=[result_path,img_path,xls_path,log_path,temp_path]
        [os.makedirs(path) for path in self.mypath if not os.path.exists(path)]
        scan_img = temp_path+r'\scan_img.png'
        self.mypath.append(scan_img)

    @debug
    def __open_browser(self):
        # 打开浏览器
        # self.browser.set_window_size(1366,760)
        self.browser.set_window_size(2560,1440)
        self.browser.get(self.sf_Express_url)
        self.__wait()
        js2='window.open("{}");'.format(self.kuaidi100_url)
        self.browser.execute_script(js2)
        self.__wait()
        all_handles = self.browser.window_handles
        self.sf_tab=all_handles[0]
        self.kd100_tab=all_handles[1]
        self.browser.switch_to_window(self.kd100_tab)
        # js1=r"""
        # var size = {}; 
        # document.body.style.zoom = size;
        # document.body.style.cssText += '; -moz-transform: scale(' + size + ');-moz-transform-origin: 0 0; ';  
        # """.format(str(self.html_size))
        # self.browser.execute_script(js1)
        self.browser.switch_to_window(self.sf_tab)

    def __scan_login(self):
        # 扫码登陆
        self.__wait()
        try:
            self.browser.find_element_by_xpath('//span[@class="agreeCookie"]').click()
        except:
            pass
        self.__wait()
        self.browser.find_element_by_xpath('//a[@class="topa maidian"]').click()
        self.__is_visible('//img[@class="scan-img"]')
        self.__wait()
        self.browser.save_screenshot(self.mypath[5])
        self.__wait()
        # 创建一个新浏览器打开扫码界面
        driver = webdriver.Firefox(executable_path=self.executable_path)
        scan_pic='file:///' + self.mypath[5].replace('\\','/')
        driver.maximize_window()
        driver.get(scan_pic)
        scan = self.__is_not_visible('//div[@class="layui-layer-shade"]')
        if scan == False:
            self.logger.info('超出60秒未进行微信扫码，直接进行查询')
            try:
                self.browser.find_element_by_xpath('//span[@class="layui-layer-setwin"]/a').click()
                self.__wait()
            except Exception:
                self.logger_error.exception("Recording Errors")
                self.logger.info('没有查询到相关物流信息')
        else:
            self.logger.info('扫码成功，开始查询')

    @debug
    def input_bill_number(self):
        # 输入快递单号开始查询
        self.logger.info('正在输入单号进行查询')
        input_box=self.browser.find_element_by_xpath('//input[@class="token-input"]')
        input_box.send_keys(self.bill_number+',')
        self.__wait()
        self.browser.find_element_by_xpath('//span[@id="queryBill"]').click()
        self.__wait()

    @debug
    def switch_to_iframe(self):
        # 切换至验证码弹窗网页框架
        self.__is_visible('//iframe')
        # 找到“嵌套”的iframe
        iframe = self.browser.find_element_by_xpath('//iframe')
        self.browser.switch_to.frame(iframe)
        self.__wait()

    @debug
    def get_verify_img(self,verify_times):
        # 获取验证码图片
        self.logger.info('第 {} 次尝试解析滑动验证码'.format(str(verify_times)))
        self.incomplete_img = self.mypath[4]+r'\{}_img_incomplete.jpg'.format(str(self.num))
        self.complete_img = self.mypath[4]+r'\{}_img_complete.jpg'.format(str(self.num))
        # 保存带缺口的滑动图片
        self.__wait()
        img = self.browser.find_element_by_xpath('//img[@id="slideBkg"]')
        incomplete_img_url = img.get_attribute('src')
        urlretrieve(url=incomplete_img_url, filename=self.incomplete_img)
        # 保存完整的滑动图片
        complete_img_url = incomplete_img_url[:-1]+'0'
        urlretrieve(url=complete_img_url, filename=self.complete_img)

    def __is_pixel_equal(self, bg_image, fullbg_image, x, y):
        # 内部函数：用于判断两张图片的像素点差异
        # 获取缺口图片的像素点(按照RGB格式)
        bg_pixel = bg_image.load()[x, y]
        # 获取完整图片的像素点(按照RGB格式)
        fullbg_pixel = fullbg_image.load()[x, y]
        # 设置一个判定值，像素值之差超过判定值则认为该像素不相同
        threshold = 50
        # 判断像素的各个颜色之差，abs()用于取绝对值
        se1 = abs(bg_pixel[0] - fullbg_pixel[0])
        se2 = abs(bg_pixel[1] - fullbg_pixel[1])
        se3 = abs(bg_pixel[2] - fullbg_pixel[2])
        if se1 < threshold and se2 < threshold and se3 < threshold:
            return True
        else:
            return False      

    def get_distance(self):
        # 获取滑动验证码需要滑动的距离
        initial_pos = 24
        fullbg_image = Image.open(self.complete_img)
        bg_image = Image.open(self.incomplete_img)
        for i in range(initial_pos, fullbg_image.size[0]):
            # 遍历像素点纵坐标
            for j in range(fullbg_image.size[1]):
                # 如果不是相同像素
                if not self.__is_pixel_equal(fullbg_image, bg_image, i, j):
                    # 返回此时横轴坐标就是滑块需要移动的距离
                    mov_pos = i/2-initial_pos
                    return mov_pos

    def get_tracks(self,mov_pos):
        # 获取滑动路径列表
        forward_tracks=[]
        mov_pos += 6  # 要回退的像素
        v0, s, t = 0, 0, 0.4  # 初速度为v0，s是已经走的路程，t是时间
        mid = mov_pos*3/5  # mid是进行减速的路程
        while s < mov_pos:
            if s < mid:  # 加速区
                a = 5
            else:  # 减速区
                a = -3
            v = v0
            tance = v*t+0.5*a*(t**2)
            tance = round(tance)
            s += tance
            v0 = v+a*t
            forward_tracks.append(tance)
        # 因为回退20像素，所以可以手动打出，只要和为20即可
        backward_tracks = [-1, -2, -1, -2]
        random.shuffle(backward_tracks)  # 洗牌
        return forward_tracks,backward_tracks

    def move_slider(self,forward_tracks,backward_tracks):
        # 开始模拟滑动验证码，先前进再后退
        self.logger.info('解析成功，正在模拟滑动验证码···')
        slider = self.browser.find_element_by_xpath(
            '//div[@id="tcaptcha_drag_button"]')
        # 使用click_and_hold()方法悬停在滑块上，perform()方法用于执行
        ActionChains(self.browser).click_and_hold(slider).perform()
        # 使用move_by_offset()方法拖动滑块，perform()方法用于执行
        for forword_track in forward_tracks:
            ActionChains(self.browser).move_by_offset(
                xoffset=forword_track, yoffset=0).perform()
        self.__wait()
        for backward_tracks in backward_tracks:
            ActionChains(self.browser).move_by_offset(
                xoffset=backward_tracks, yoffset=0).perform()
        # 模拟人类对准时间
        self.__wait()
        # 释放滑块
        ActionChains(self.browser).release().perform()
        try:
            if self.__is_not_visible('//p[@class="tcaptcha-title"]',10):
                verify_ok = True
            else:
                verify_ok = False
        except Exception:
            verify_ok = True
        finally:
            return verify_ok

    @debug
    def switch_to_default(self):
        # 切换至主页面
        self.browser.switch_to.default_content()
        self.__wait()

    def __get_route_pic(self,roll=0,pic_n=''):
        js='document.getElementsByClassName("routes-wrapper")[0].scrollTop='+str(roll)
        self.browser.execute_script(js)
        self.browser.save_screenshot(self.screenshot_img_temp.replace('.png',pic_n))
        return Image.open(self.screenshot_img_temp.replace('.png',pic_n))

    @debug
    def get_route_info(self):
        # 获取物流节点信息和截图
        self.route = ''
        self.logger.info('获取物流节点信息和截图')
        self.screenshot_img1 = self.mypath[1] + r'\%s-WLJD-%s.png' % (self.shot_name, self.bill_number)
        self.screenshot_img_temp = self.mypath[4] + r'\%s-WLJD-%s.png' % (self.shot_name, self.bill_number)
        try:
            delivery_map=self.browser.find_element_by_xpath('//div[@class="delivery"]/div[1]').get_attribute('class')
            target=self.browser.find_element_by_xpath('//input[@type="checkbox"]')
            if delivery_map=='delivery-item send-out-item' or delivery_map=='delivery-item send-out-item ' :
                js = "document.getElementsByClassName('delivery-item send-out-item')[0].className = 'delivery-item send-out-item  brief-model'"
                self.browser.execute_script(js)
            # delivery-item send-out-item  brief-model
            self.__wait()
            self.browser.execute_script("arguments[0].scrollIntoView();", target)
            self.__wait()
            routes = self.browser.find_elements_by_xpath('//div[@class="route-list"]/ul')
            for rou in routes:
                r = rou.text.replace('\n', ' ')
                self.logger.info(r)
                self.route = self.route+'\n'+ r
            if 0< len(routes) <=8:
                newpic=self.__get_route_pic(0,'.png')
            elif 8 < len(routes)<=16:
                im1=self.__get_route_pic(0,'_1.png')
                im2=self.__get_route_pic(285,'_2.png')
                self.logger.info('正在拼接物流节点截图')
                x,y=im1.size[0],im1.size[1]
                sb,xb=290,(1344-580) #上边距,下边距
                y1=im1.size[1]-sb-xb
                im1 = im1.crop((0, 0, x, y-xb))
                im2 = im2.crop((0, sb, x, y-xb))
                newpic = Image.new('RGB', (x, y+y1-xb))  
                newpic.paste(im1, (0, 0, x, y-xb))
                newpic.paste(im2, (0, y-xb, x, y-xb+y1))
                newpic.save(self.screenshot_img1)
            elif len(routes)>16:
                im1=self.__get_route_pic(0,'_1.png')
                im2=self.__get_route_pic(285,'_2.png')
                im3=self.__get_route_pic(570,'_3.png')  
                self.logger.info('正在拼接物流节点截图')
                x,y=im1.size[0],im1.size[1]
                sb,xb=290,(1344-580) #上边距,下边距
                y1=im1.size[1]-sb-xb
                im1 = im1.crop((0, 0, x, y-xb))
                im2 = im2.crop((0, sb, x, y-xb))
                im3 = im3.crop((0, sb, x, y-xb))
                newpic = Image.new('RGB', (x, y+y1*2-xb))  
                newpic.paste(im1, (0, 0, x, y-xb))
                newpic.paste(im2, (0, y-xb, x, y-xb+y1))
                newpic.paste(im3, (0, y-xb+y1, x, y-xb+y1+y1))
                newpic.save(self.screenshot_img1)
            newpic = newpic.crop((790, 90, newpic.size[0]-790, newpic.size[1]))
            newpic.save(self.screenshot_img1)

        except Exception:
            self.logger.info('没有查询到相关物流信息')

    def get_dzcg_info(self):
        # 获取电子存根信息和截图
        self.logger.info('正在获取电子存根信息和截图')
        self.screenshot_img2 = self.mypath[1] + r'\%s-YDXQ-%s.png' % (self.shot_name, self.bill_number)
        self.screenshot_img3 = self.mypath[1] + r'\%s-DZCG-%s.png' % (self.shot_name, self.bill_number)
        self.dzcg = ''
        try:
            self.browser.find_element_by_xpath(
                '//div[@class="operates-wrapper"]/a').click()
            self.__wait()
            target = self.browser.find_element_by_xpath(
                '//a[@aria-controls="waybillDetail"]')
            target.click()
            self.browser.execute_script(
                "arguments[0].scrollIntoView();", target)
            self.__wait()
            self.browser.save_screenshot(self.screenshot_img2)
            newpic=Image.open(self.screenshot_img2)
            newpic = newpic.crop((980, 460, newpic.size[0]-995, newpic.size[1]-525))
            newpic.save(self.screenshot_img2)
            target = self.browser.find_element_by_xpath(
                '//a[@aria-controls="electronicStub"]')
            target.click()
            self.browser.execute_script(
                "arguments[0].scrollIntoView();", target)
            self.__wait()
            self.browser.save_screenshot(self.screenshot_img3)
            newpic=Image.open(self.screenshot_img3)
            newpic = newpic.crop((980, 190, newpic.size[0]-990, newpic.size[1]-520))
            newpic.save(self.screenshot_img3)
            dzcgs = self.browser.find_elements_by_xpath(
                '//table[@class="borderrb"]/tbody/tr')
            for dz in dzcgs:
                r = dz.text.replace('\n', ' ')
                self.logger.info(r)
                self.dzcg = self.dzcg+'\n'+r
        except Exception:
            self.logger.info('该单号对应的手机号与扫码的微信号可能不符，无法获取电子存根截图')

    @debug
    def next_bill_number(self):
        # 关闭电子存根界面，开始查下一个单号
        try:
            self.browser.find_element_by_xpath(
                '//button[@class="close"]').click()
            self.__wait()
        except Exception:
            self.logger_error.exception("Recording Errors")
        self.browser.find_element_by_xpath(
            '//a[@class="close"]').click()
        self.browser.find_element_by_xpath(
            '//span[@class="sfi sfi-trash"]').click()
        self.__wait()

    def sf_single(self):
        ''' 顺丰主体 '''
        self.browser.switch_to_window(self.sf_tab)
        self.__wait()
        self.input_bill_number()
        self.switch_to_iframe()
        verify_times,verify_ok=0,False
        while verify_times<=4 and verify_ok==False:
            try:
                verify_times+=1
                self.browser.find_element_by_xpath('//div[@id="reload"]').click()
                self.get_verify_img(verify_times)
                mov_pos=self.get_distance() 
                forward_tracks,backward_tracks=self.get_tracks(mov_pos)  
                verify_ok=self.move_slider(forward_tracks,backward_tracks)  
            except Exception:
                self.logger_error.exception("Recording Errors")
                self.__wait()
            finally:
                self.logger.info('verify_ok = ' + str(verify_ok))
        if verify_times > 4:
            self.logger.info('验证解析尝试5次失败，不再尝试，开始下一个单号')
        if verify_ok==True:
            self.switch_to_default()
            self.get_route_info()
            self.get_dzcg_info()
            self.next_bill_number()

    def kd100_single(self):
        self.browser.switch_to_window(self.kd100_tab)
        self.__wait()
        try:
            self.browser.find_element_by_xpath('//span[@class="coupon-pop-close"]').click()
        except:
            pass
        self.__wait()
        self.__wait()
        input_box=self.browser.find_element_by_xpath('//input[@id="postid"]')
        self.browser.execute_script("arguments[0].scrollIntoView();",input_box)
        input_box.send_keys(self.bill_number) 
        self.__wait()
        self.browser.find_element_by_xpath('//a[@id="query"]').click()
        self.__wait()
        # 物流节点
        self.logger.info('正在获取物流节点详细信息')
        self.dzcg = ''
        self.route= ''
        try:
            routes = self.browser.find_elements_by_xpath(
                '//table[@class="result-info"]/tbody/tr')
            for rou in routes:
                r = rou.text.replace('\n', ' ')
                self.logger.info(r)
                self.route = self.route+'\n'+r
        except Exception:
            self.logger_error.exception("Recording Errors")
            self.logger.info('没有查询到相关物流信息')
        self.logger.info('正在获取物流节点截图')    
        notFindTip=self.browser.find_elements_by_xpath('//div[@id="notFindTip"]')[0].get_attribute("style") 
        # 获取截图
        if notFindTip=='display: none;':
            self.screenshot_img4=self.mypath[1] + r'\%s-WLJD-%s.png' % (self.num, self.bill_number)
            self.browser.save_screenshot(self.screenshot_img4)
            self.__wait()
            im = Image.open(self.screenshot_img4)
            im1 = im.crop((800, 0, im.size[0]-800, im.size[1]-100))
            im1.save(self.screenshot_img4)
            self.logger.info('物流节点截图已保存') 
        else:
            self.logger.info('该单号没有物流信息,暂不截图')

    def get_comcode(self):
        if self.bill_number[:2]=="SF":
            code = 'shunfeng'
        else:
            headers = {'User-Agent': random.choice(self.user_agent_list),
                            'Host': 'www.kuaidi100.com',
                            'Connection': 'keep-alive',
                            'Content-Length': '0',
                            'Accept': 'application/json, text/javascript, */*; q=0.01',
                            'X-Requested-With': 'XMLHttpRequest',
                            'Sec-Fetch-Mode': 'cors',
                            'Origin': 'https://www.kuaidi100.com',
                            'Sec-Fetch-Site': 'same-origin',
                            'Referer': 'https://www.kuaidi100.com/',
                            'Accept-Encoding': 'gzip, deflate, br',
                            'Accept-Language': 'zh-CN,zh;q=0.9',
                            }
            url = 'https://www.kuaidi100.com/autonumber/autoComNum?resultv2=1&text='+self.bill_number
            data = {'resultv2': '1', 'text': self.bill_number}
            try:
                r = requests.post(url, headers=headers, data=data)
                result = r.json()
                code_en = [i['comCode'] for i in result['auto']]
                code=code_en[0]
            except:
                code = 'unknow'
        
        return code

    def main(self,shot_name,bill_number=[]):
        try:
            self.num,self.bill_number=bill_number[0],str(bill_number[1])
            self.shot_name=shot_name
            self.logger.info('当前进度：{}/{} ，正在查询的单号为 [{}]'.format(self.num,self.count,self.bill_number))
            comcode=self.get_comcode()
            if comcode=='unknow':
                self.logger.info('单号 [{}] 所属快递公司未知，无法查询'.format(self.bill_number))
            else:
                self.logger.info('单号 [{}] 所属快递公司为 [{}]'.format(self.bill_number,comcode))
                if comcode=='shunfeng':
                    self.sf_single()
                else:
                    self.kd100_single()
        except Exception:
            self.logger_error.exception("Recording Errors")
        finally:
            data=[self.bill_number,self.route,self.dzcg]
            print('\n')
            return data

    def __del__(self):
        self.browser.quit()


class Input_and_Output(object):

    def __validateFileName(self,FileName):
        rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
        new_FileName = re.sub(rstr, "-", FileName)  # 替换为下划线
        return new_FileName

    def load_input(self):
        try:
            input_path=os.path.abspath(os.curdir)+r'\Input\bill_number.xlsx'
            df = pd.read_excel(input_path).dropna()
            bill_number_list=(df["2"]).tolist()[1:]
            shot_name_list=(df["3"]).tolist()[1:]
            shot_name_list = [self.__validateFileName(i) for i in shot_name_list]
            if len(bill_number_list) == 0:
                print('bill_number.xlsx 中未发现快递单号')      
            return bill_number_list,shot_name_list
        except Exception as e:
            print('快递单号读取错误，检查是否存在 bill_number.xlsx 工作簿') 
        
    def output_csv(self, output_path='',data=[]):
        '''输出文件'''
        self.output_path=output_path + r'\express_route.csv'
        with open(self.output_path, 'a', newline='', encoding='utf-8-sig')as f:
            writer = csv.writer(f)
            writer.writerow(data)

    def csv_to_xlsx(self):
        '''转换为xlsx'''
        csv = pd.read_csv(self.output_path, encoding='utf-8-sig')
        csv.to_excel(self.output_path.replace('.csv', '.xlsx'), sheet_name='15Seconds')


def main():
    start=time.time()
    io=Input_and_Output() # 实例1
    bill_number_list,shot_name_list=io.load_input()
    count=len(bill_number_list)
    bill_number_list=list(enumerate(bill_number_list,start=1))
    
    if count>0:
        track=Express_Tracking(count)# 实例2
        mypath=track.mypath
        title=['物流单号','物流节点','电子存根']
        io.output_csv(mypath[2],title)
        for i in range(len(bill_number_list)):
            data = track.main(shot_name_list[i],bill_number_list[i])
            io.output_csv(mypath[2],data)
        io.csv_to_xlsx()
        if count/100==0:
            track.logger.info('每查询100个，休息2分钟')
            time.sleep(120)
    else:
        pass
    stamp = time.time() - start
    mon=round(stamp/60)
    track.logger.info('所有快递单号查询结束，总用时约 {} 分钟'.format(mon))
    del track

if __name__ == "__main__":
    main()
    # io=Input_and_Output() # 实例1
    # bill_number_list=io.load_input()
    # print(bill_number_list)
