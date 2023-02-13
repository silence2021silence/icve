# -*- coding=utf-8 -*-
"""
Time:        2023/02/13 21:00
Version:     V 0.0.1
File:        icve.py
Describe:    
Author:      Lanyu
E-Mail:      silence2021silence@163.com
Github link: https://github.com/silence2021silence/
Gitee link:  https://gitee.com/silence2021silence/
"""

# -*- coding=utf-8 -*-

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from lxml import etree

print("本软件仅供学习研究娱乐用,请遵守相关规定,本软件已开源,禁止用于商业用途\n")
print("作者:蓝鱼Lanyu,使用教程或更新或问题反馈请关注博客网址:geeklanyu.com\n")
course_id = input("输入params.courseId\n")
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9999")
driver = webdriver.Chrome(options=options)

# 解析html
html = open("data.html", "r", encoding="utf-8").read()
tree = etree.HTML(html)

# 获取所有视频项目的xpath对象
all_video_obj = tree.xpath("//div[@class='s_learn_type s_learn_video']/..")

# 筛选出已完成和未完成的视频
finished_video_id = []
unfinished_video_id = []
for i in all_video_obj:
    t = i.xpath("div[1]/@class='item_done_icon item_done_pos'")
    if t:
        unfinished_video_id.append(i.xpath("./@id")[0].replace("s_point_", ""))
    else:
        finished_video_id.append(i.xpath("./@id")[0].replace("s_point_", ""))

# 嵌入参数生成请求地址
urls = []
for i in unfinished_video_id:
    url = "https://course.icve.com.cn/learnspace/learn/learn/templateeight/content_video.action?params.courseId=%s&params.itemId=%s&params.templateStyleType=0&_t=%d" % (course_id, i, round(time.time() * 1000))
    urls.append(url)


# 获取视频长度
def get_length():
    while True:
        html = driver.page_source
        tree = etree.HTML(html)
        length_str = tree.xpath("//span[@id='screen_player_time_2']//text()")[0]
        h_m_s = [int(i) for i in length_str.split(":")]
        # 转换为秒
        if len(h_m_s) == 3:
            s = int(h_m_s[0]) * 3600 + int(h_m_s[1]) * 60 + int(h_m_s[2])
            if s != 0:
                return s
        elif len(h_m_s) == 2:
            s = int(h_m_s[0]) * 60 + int(h_m_s[1])
            if s != 0:
                return s


# 发出请求
for i in urls:
    driver.get(i)
    driver.find_element(By.XPATH, value="*").send_keys(Keys.SPACE)
    time.sleep(get_length() + 5)

