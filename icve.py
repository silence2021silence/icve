# -*- coding=utf-8 -*-
"""
Time:        2023/2/26 17:00
Version:     V 0.0.3
File:        icve-v0.0.3.py
Describe:    
Author:      Lanyu
E-Mail:      silence2021silence@163.com
Github link: https://github.com/silence2021silence/
Gitee link:  https://gitee.com/silence2021silence/
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from lxml import etree


def run1():
    driver = webdriver.Chrome(executable_path="chromedriver.exe")
    driver.get("https://icve-mooc.icve.com.cn/cms/")
    print("谷歌浏览器连接成功")
    input("登录账号并进入课程学习后按回车键\n")
    run2(driver)


def run2(driver):
    item_type = input("\n[1]视频\n[2]音频\n[3]文档\n[4]图文\n选择哦~\n")
    print("好的~任务开始")

    # 解析html
    try:
        html = open("data.html", "r", encoding="utf-8").read()
    except FileNotFoundError:
        print("找不到data.html哦")
        run1()
    else:
        tree = etree.HTML(html)
        course_id = tree.xpath("//input[@id='currentCourseId']/@value")[0]
        if item_type == "1":
            item_type = "video"
            all_video_obj = tree.xpath("//div[@class='s_learn_type s_learn_video']/..")
            unfinished_ids, unfinished_names = get_unfinished_item(all_video_obj)
            urls = make_urls(unfinished_ids, item_type, course_id)
            send_request(urls, driver, item_type, unfinished_names)

        elif item_type == "2":
            item_type = "audio"
            all_audio_obj = tree.xpath("//div[@class='s_learn_type s_learn_audio']/..")
            unfinished_ids, unfinished_names = get_unfinished_item(all_audio_obj)
            urls = make_urls(unfinished_ids, item_type, course_id)
            send_request(urls, driver, item_type, unfinished_names)

        elif item_type == "3":
            item_type = "doc"
            all_doc_obj = tree.xpath("//div[@class='s_learn_type s_learn_doc']/..")
            unfinished_ids, unfinished_names = get_unfinished_item(all_doc_obj)
            urls = make_urls(unfinished_ids, item_type, course_id)
            send_request(urls, driver, item_type, unfinished_names)

        elif item_type == "4":
            item_type = "text"
            all_doc_obj = tree.xpath("//div[@class='s_learn_type s_learn_text']/..")
            unfinished_ids, unfinished_names = get_unfinished_item(all_doc_obj)
            urls = make_urls(unfinished_ids, item_type, course_id)
            send_request(urls, driver, item_type, unfinished_names)

        else:
            print("输入错误")
            run1()


# 爬取未完成的项目
def get_unfinished_item(obj):
    print("正在爬取未完成的项目")
    unfinished_ids = []
    unfinished_names = []
    for i in obj:
        t = i.xpath("div[1]/@class='item_done_icon item_done_pos'")
        if t:
            unfinished_ids.append(i.xpath("./@id")[0].replace("s_point_", ""))
            unfinished_names.append(i.xpath("div[3]//text()")[0])
    print("爬取完毕")
    return unfinished_ids, unfinished_names


# 嵌入参数生成请求地址
def make_urls(item_ids, item_type, course_id):
    print("正在生成请求地址")
    urls = []
    for i in item_ids:
        url = "https://course.icve.com.cn/learnspace/learn/learn/templateeight/content_%s.action?params.courseId=%s&params.itemId=%s&params.templateStyleType=0&_t=%d" % (
            item_type, course_id, i, round(time.time() * 1000))
        urls.append(url)
    print("生成完毕")
    return urls


# 发送请求
def send_request(urls, driver, item_type, unfinished_names):
    for i in range(len(urls)):
        print("开始处理", unfinished_names[i], "[%d/%d]" % (unfinished_names.index(unfinished_names[i]) + 1, len(unfinished_names)))
        print("发送请求")
        driver.get(urls[i])
        windows = driver.window_handles
        driver.switch_to.window(windows[0])
        print("请求成功")
        time.sleep(1)
        driver.find_element(By.XPATH, value="*").send_keys(Keys.SPACE)
        time.sleep(get_length(driver, item_type) + 5)
        print("等待播放完毕")
        print("处理完毕")
    input("任务结束~如要继续执行任务按回车键")
    run2(driver)


# 获取视频/音频长度
def get_length(driver, item_type):
    xpath = ""
    if item_type == "video":
        xpath = "//span[@id='screen_player_time_2']//text()"
    if item_type == "audio":
        xpath = "//div[@class='audio-time audio-time-duration']//text()"
    if item_type == "doc" or item_type == "text":
        return 0

    print("正在获取视频/音频长度")
    while True:
        html = driver.page_source
        tree = etree.HTML(html)
        length_str = tree.xpath(xpath)[0]
        h_m_s = [int(i) for i in length_str.split(":")]
        # 转换为秒
        if len(h_m_s) == 3:
            s = int(h_m_s[0]) * 3600 + int(h_m_s[1]) * 60 + int(h_m_s[2])
            if s != 0:
                print("获取完毕")
                return s
        elif len(h_m_s) == 2:
            s = int(h_m_s[0]) * 60 + int(h_m_s[1])
            if s != 0:
                print("获取完毕")
                return s


if __name__ == "__main__":
    print("hello哇~")
    time.sleep(0.5)
    print("我是本软件的作者[Lanyu(蓝鱼)]")
    time.sleep(0.5)
    print("使用教程|检查更新|问题反馈 前往我的博客[geeklanyu.com]查看哦")
    time.sleep(0.5)
    print("本软件仅供学习/研究/娱乐用,请遵守相关规定,本软件已开源,禁止用于商业用途")
    time.sleep(0.5)
    run1()
