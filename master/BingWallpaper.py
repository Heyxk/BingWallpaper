# coding=utf-8
import ConfigParser
import codecs
import logging
import os
import re
import win32api

import sys
import win32con
import win32gui
import time
import requests
from PIL import Image


class BingWallpaper(object):
    def __init__(self):
        """
        初始数据

        """
        # 请求头
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
        # 利用os.path.abspath()返回格式化路径，防止出错
        # 用户目录下的图片文件夹
        self.save_path = os.path.normpath(os.path.expanduser('~') + '\\Pictures\\BingWallpaper')
        self.day = '0'  # 默认获取当天的图片
        self.startup = False

    def makedir(self, img_dir):  # 此函数创建文件夹，并移动到此文件夹下
        try:
            # root_dir = os.path.expanduser('~') + '\\Pictures'  # 用户目录下的图片文件夹
            # img_dir = self.default_path  # 图片存放路径
            isExists = os.path.exists(img_dir)  # 判断文件夹是否存在
            if not isExists:
                os.makedirs(img_dir)  # 创建文件夹
            return img_dir
        except WindowsError:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return os.path.dirname(os.path.normpath(sys.argv[0]))  # 返回文件真实路径

    def download_img(self, day):
        try:
            index_url = 'http://www.bing.com/HPImageArchive.aspx?format=js&idx=%s&n=1&mkt=en-US' % day
            res = requests.get(index_url, headers=self.headers, timeout=(3, 3))
            img_info = res.json()["images"][0]
            # tooltipsInfo = res.json()["tooltips"]
            img_url = 'http://cn.bing.com' + img_info["url"]
            title = img_info["startdate"] + ' ' + re.sub(r'.\(.*', '', img_info["copyright"])
            img_name = self.save_path + '\\' + title + '.jpg'
            with open(img_name, 'wb') as img_obj:
                img_content = requests.get(img_url, headers=self.headers, timeout=(3, 3)).content  # 获取图片的二进制数据
                img_obj.write(img_content)
                img_obj.flush()
            return img_name
        except:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return None

    def set_wallpaper_from_bmp(self, bmp_path):
        try:
            # 打开指定注册表路径
            hKey = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0,
                                         win32con.KEY_SET_VALUE)
            # 最后的参数:2拉伸,0居中,6适应,10填充,0平铺
            win32api.RegSetValueEx(hKey, "WallpaperStyle", 0, win32con.REG_SZ, "2")
            # 最后的参数:1表示平铺,拉伸居中等都是0
            win32api.RegSetValueEx(hKey, "TileWallpaper", 0, win32con.REG_SZ, "0")
            # 刷新桌面
            win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, bmp_path, win32con.SPIF_SENDWININICHANGE)
            win32api.RegCloseKey(hKey)
            return True
        except:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False

    def set_wallpaper(self, img_path):
        try:
            # 把图片格式统一转换成bmp格式,并放在源图片的同一目录
            img_dir = os.path.dirname(img_path)
            new_bmp_path = os.path.join(img_dir, 'now_wallpaper.bmp')
            bmp_image = Image.open(img_path)
            bmp_image.save(new_bmp_path, "BMP")
            self.set_wallpaper_from_bmp(new_bmp_path)
            return True
        except:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False

    def load_config(self):
        # 读取配置文件
        try:
            # 配置文件不存在则调用函数创建，并直接返回
            if not os.path.exists('config.conf'):
                self.init_config()
                return True
            # 若存在配置文件则读取
            cp = ConfigParser.SafeConfigParser(allow_no_value=True)
            # cp.read('config.conf')  # 方式一
            with codecs.open('config.conf', 'r', encoding='utf-8') as f:  # 有Unicode字符时用此种方式
                cp.readfp(f)

            if cp.has_section('CUSTOM'):
                if cp.has_option('CUSTOM', 'value') and not cp.getboolean('CUSTOM', 'value'):
                    return False
                if cp.has_option('CUSTOM', 'customfolder') and cp.get('CUSTOM', 'customfolder'):
                    self.save_path = os.path.abspath(cp.get('CUSTOM', 'customfolder'))  # 格式化路径
                if cp.has_option('CUSTOM', 'startup') and cp.get('CUSTOM', 'startup'):  # 有键且存在值
                    self.startup = cp.getboolean('CUSTOM', 'startup')
                if cp.has_option('CUSTOM', 'day') and cp.get('CUSTOM', 'day'):
                    if int(cp.get('CUSTOM', 'day')) < -1:
                        return False
                    self.day = cp.get('CUSTOM', 'day')
            else:
                return False
        except ConfigParser.Error:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False
        except IOError:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False
        except ValueError:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False

    def init_config(self):
        try:
            cp = ConfigParser.SafeConfigParser(allow_no_value=True)
            cp.add_section('CUSTOM')
            cp.set('CUSTOM', 'value', 'True')
            cp.set('CUSTOM', 'customfolder', self.save_path)
            cp.set('CUSTOM', 'day', self.day)
            cp.set('CUSTOM', 'startup', 'True')
            cp.write(open('config.conf', 'w'))
            return True
        except:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            return False

    def add2boot(self, startup):
        try:
            # "注册到启动项"
            path = os.path.normpath(sys.argv[0])
            subKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
            hKey = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, subKey, 0, win32con.KEY_ALL_ACCESS)
            # win32api.RegQueryValueEx(hKey, 'BingWallpaper')  # 判断项值是否存在
            if startup:
                try:
                    win32api.RegQueryValueEx(hKey, 'BingWallpaper')  # 判断项值是否存在
                except:
                    win32api.RegSetValueEx(hKey, 'BingWallpaper', 0, win32con.REG_SZ, path)
            else:
                try:
                    win32api.RegDeleteValue(hKey, 'BingWallpaper')
                except:
                    pass
        except:
            logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        finally:
            win32api.RegCloseKey(hKey)


if __name__ == "__main__":
    try:
        bing = BingWallpaper()
        # 创建日志文件
        logging.basicConfig(filename='BingWallpaper.log')
        # 加载配置
        bing.load_config()
        bing.add2boot(bing.startup)
        # 创建文件夹
        bing.save_path = bing.makedir(bing.save_path)
        bing.img_path = bing.download_img(bing.day)
        # 判断图片是否下载成功
        if bing.img_path:
            bing.set_wallpaper(bing.img_path)
    except:
        logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
