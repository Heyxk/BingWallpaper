# coding=utf-8
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

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/52.0.2743.116 Safari/537.36 Edge/15.15063'}
save_path = os.path.normpath(os.path.expanduser('~') + '\\Pictures\\BingWallpaper')  # 用户目录下的图片文件夹
day = '0'  # 默认获取当天图片


def makedir(img_dir):  # 此函数创建文件夹，并移动到此文件夹下
    try:
        isExists = os.path.exists(img_dir)  # 判断文件夹是否存在
        if not isExists:
            os.makedirs(img_dir)  # 创建文件夹
        return img_dir
    except WindowsError:
        logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        return os.path.dirname(os.path.normpath(sys.argv[0]))  # 返回文件真实路径


def download_img(to_path, target_date):
    try:
        index_url = 'http://www.bing.com/HPImageArchive.aspx?format=js&idx=%s&n=1&mkt=en-US' % target_date
        res = requests.get(index_url, headers=headers, timeout=(3, 3))
        img_info = res.json()["images"][0]
        # tooltipsInfo = res.json()["tooltips"]
        img_url = 'http://cn.bing.com' + img_info["url"]
        title = img_info["startdate"] + ' ' + re.sub(r'.\(.*', '', img_info["copyright"])
        img_name = to_path + '\\' + title + '.jpg'
        with open(img_name, 'wb') as img_obj:
            img_content = requests.get(img_url, headers=headers, timeout=(3, 3)).content  # 获取图片的二进制数据
            img_obj.write(img_content)
            img_obj.flush()
        return img_name
    except:
        logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        return None


def set_wallpaper_from_bmp(bmp_path):
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


def set_wallpaper(_img_path):
    try:
        # 把图片格式统一转换成bmp格式,并放在源图片的同一目录
        img_dir = os.path.dirname(_img_path)
        new_bmp_path = os.path.join(img_dir, 'now_wallpaper.bmp')
        bmp_image = Image.open(_img_path)
        bmp_image.save(new_bmp_path, "BMP")
        set_wallpaper_from_bmp(new_bmp_path)
        return True
    except:
        logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        return False


if __name__ == '__main__':
    try:
        logging.basicConfig(filename='BingWallpaper.log')
        save_path = makedir(save_path)
        img_path = download_img(save_path, '0')
        if img_path:
            set_wallpaper(img_path)
    except:
        logging.exception(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
