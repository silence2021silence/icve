# -*- coding=utf-8 -*-
"""
Time:        2022/12/28 22:40
Version:     V 0.0.4
File:        icve.py
Describe:    
Author:      Lanyu
E-Mail:      silence2021silence@163.com
Github link: https://github.com/silence2021silence/
Gitee link:  https://gitee.com/silence2021silence/
"""

import time
import pyautogui
import os
import requests
import xlrd
from bs4 import BeautifulSoup


class OldMOOC:
    def get_task(self):
        workbook = xlrd.open_workbook("task.xls")
        worksheet = workbook.sheet_by_index(0)

        # 控件列表
        try:
            widget_list = ["img/" + worksheet.col(0)[1].value, "img/" + worksheet.col(0)[3].value,
                           "img/" + worksheet.col(0)[5].value, "img/" + worksheet.col(0)[7].value,
                           "img/" + worksheet.col(0)[9].value, "img/" + worksheet.col(0)[11].value,
                           "img/" + worksheet.col(0)[13].value]
        except IndexError:
            print("表格第一列填写有误")
            input("请输入任意内容后按回车键退出\n")
            exit()

        widget_list = ["img/" + worksheet.col(0)[1].value, "img/" + worksheet.col(0)[3].value,
                       "img/" + worksheet.col(0)[5].value, "img/" + worksheet.col(0)[7].value,
                       "img/" + worksheet.col(0)[9].value, "img/" + worksheet.col(0)[11].value,
                       "img/" + worksheet.col(0)[13].value]

        # 任务列表
        task_list = []
        # 统计各列的任务数的列表
        task_num = []
        for i in range(worksheet.ncols - 1):
            for j in range(len(worksheet.col(i)) - 1):
                value = worksheet.col(i + 1)[j + 1].value
                task_list.append("img/" + value)
                while "img/" in task_list:
                    task_list.remove("img/")
            task_num.append(len(task_list))

        # 将控件列表和任务列表合并成最终的任务列表
        final_task_list = []
        for i in task_list:
            final_task_list.append(i)
            for j in widget_list:
                final_task_list.append(j)

        # 删除模块到第一个任务之间的控件
        for i in range(len(task_num) - 1):
            del final_task_list[(task_num[i]) * 7 + (task_num[i]) + 1:(task_num[i]) * 7 + (task_num[i]) + 8]
        del final_task_list[1:8]

        return final_task_list, widget_list


    def click(self, img, widget_list):
        verification = widget_list[1]
        verification_bar_0 = widget_list[2]
        verification_bar_1 = widget_list[3]
        replay_key = widget_list[4]
        while True:
            if not os.path.exists(img):
                print("找不到" + img + "文件")
                input("请输入任意内容后按回车键退出\n")
                exit()
            time.sleep(1)
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if img == verification and location is not None:
                location_vb0 = pyautogui.locateCenterOnScreen(verification_bar_0, confidence=0.9)
                location_vb1 = pyautogui.locateCenterOnScreen(verification_bar_1, confidence=0.9)
                if location_vb0 is not None and location_vb1 is not None:
                    pyautogui.moveTo(location_vb0.x, location_vb0.y, duration=0.5)
                    pyautogui.dragTo(location_vb1.x, location_vb1.y, duration=0.5, button="left")
                    break
            elif img == verification:
                break
            elif img == verification_bar_0:
                break
            elif img == verification_bar_1:
                break
            elif img != verification and img != verification_bar_0 and img != verification_bar_1 and location is not None:
                pyautogui.click(location.x, location.y, interval=0.2, duration=0.5, button="left")
                break
            else:
                if img == replay_key:
                    print("等待播放完毕")
                    time.sleep(1)
                else:
                    time.sleep(1)
                    pyautogui.scroll(-100)
                    print("寻找下一个任务")


    def main(self):
        print("程序已启动，2秒后开始操作")
        time.sleep(2)
        if not os.path.exists("img"):
            print("找不到img文件夹")
            input("请输入任意内容后按回车键退出\n")
            exit()
        if not os.path.exists("task.xls"):
            print("找不到task.xls文件")
            input("请输入任意内容后按回车键退出\n")
            exit()
        final_task_list, widget_list = self.get_task()
        i = 0
        while i < len(final_task_list):
            img = final_task_list[i]
            self.click(img, widget_list)
            print("点击", img)
            i += 1
        input("所有任务执行完毕，请输入任意内容后按回车键退出\n")
        exit()


class NewMOOC:
    def get_task(self):
        workbook = xlrd.open_workbook("task_new.xls")
        worksheet = workbook.sheet_by_index(0)

        # 控件列表
        try:
            widget_list = ["img/" + worksheet.col(0)[1].value, "img/" + worksheet.col(0)[3].value]
        except IndexError:
            print("表格第一列填写有误")
            input("请输入任意内容后按回车键退出\n")
            exit()

        widget_list = ["img/" + worksheet.col(0)[1].value, "img/" + worksheet.col(0)[3].value]

        # 任务列表
        task_list1 = []
        # 统计各列的任务数的列表
        task_num = []
        for i in range(worksheet.ncols - 1):
            for j in range(len(worksheet.col(i)) - 1):
                value = worksheet.col(i + 1)[j + 1].value
                task_list1.append("img/" + value)
                while "img/" in task_list1:
                    task_list1.remove("img/")
            task_num.append(len(task_list1))

        # 将控件列表和任务列表合并
        task_list2 = []
        for i in task_list1:
            task_list2.append(i)
            for j in widget_list:
                task_list2.append(j)

        # 删除模块到第一个任务之间的控件
        for i in range(len(task_num) - 1):
            del task_list2[(task_num[i]) * 2 + (task_num[i]) + 1:(task_num[i]) * 2 + (task_num[i]) + 3]
        del task_list2[1:3]

        # 只含有任务的列表
        task_list3 = []
        for i in range(len(task_num)):
            if task_num[i] == 2:
                value = worksheet.col(i + 1)[2].value
                task_list3.append("img/" + value)
            elif task_num[i] > 2:
                values = worksheet.col(i + 1)[2:task_num[i]]
                for j in values:
                    task_list3.append("img/" + j.value)
            else:
                print("表格第%d列填写有误" % (i + 1))
                input("请输入任意内容后按回车键退出\n")
                exit()

        # 只含有项目的列表一式两份
        task_list4 = []
        task_list5 = []
        for i in range(len(task_num)):
            for j in range(task_num[i] - 1):
                value = worksheet.col(i + 1)[1].value
                task_list4.append("img/" + value)
                task_list5.append("img/" + value)

        # 添加折叠项防止误点
        task_list6 = []
        for i in task_list2:
            if i == task_list2[3]:
                task_list6.append(i)
                task_list6.append(task_list3[0])
                del task_list3[0]
                task_list6.append(task_list4[0])
                del task_list4[0]
                task_list6.append(task_list5[0])
                del task_list5[0]
            else:
                task_list6.append(i)

        return task_list6, widget_list

    def click(self, img, widget_list):
        replay_key = widget_list[1]
        while True:
            if not os.path.exists(img):
                print("找不到" + img + "文件")
                input("请输入任意内容后按回车键退出\n")
                exit()
            time.sleep(1)
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, interval=0.2, duration=0.5, button="left")
                break
            else:
                if img == replay_key:
                    print("等待播放完毕")
                    time.sleep(1)
                else:
                    time.sleep(1)
                    pyautogui.moveTo(1700, 600, duration=0.5)  # 挖坑 有空再补：优化为检测屏幕分辨率 猜测右边目录窗口所占屏幕的大概比例 并将指针移至此坐标
                    pyautogui.scroll(-100)
                    print("寻找下一个任务")

    def main(self):
        print("程序已启动，2秒后开始操作")
        time.sleep(2)
        if not os.path.exists("img"):
            print("找不到img文件夹")
            input("请输入任意内容后按回车键退出\n")
            exit()
        if not os.path.exists("task_new.xls"):
            print("找不到task_new.xls文件")
            input("请输入任意内容后按回车键退出\n")
            exit()
        task_list, widget_list = self.get_task()
        print(task_list)
        i = 0
        while i < len(task_list):
            img = task_list[i]
            self.click(img, widget_list)
            print("点击", img)
            i += 1
        input("所有任务执行完毕，请输入任意内容后按回车键退出\n")
        exit()


def welcome():
    print("本程序已开源，源代码、打包可执行文件(.exe文件)、使用说明、配置教程都在GitHub和Gitee里，欢迎来Star和Fork，地址：")
    print("https://github.com/silence2021silence/")
    print("https://gitee.com/silence2021silence/")
    print("技术支持与意见反馈可直接在仓库建issues或者关注微信公众号“geeklanyu”留言或者联系邮箱“silence2021silence@163.com”")
    print("免责声明：本程序仅供学习、研究与娱乐使用，使用本程序违反相关法律或相关规章制度的与作者无关，禁止用于任何商业用途。")
    text = input("请输入“同意”或“不同意”后按回车键\n")
    if text != "同意":
        exit()
    else:
        version = "v0.0.4"
        print("当前版本为%s，正在检查更新..." % version)
        html = requests.get("https://gitee.com/silence2021silence/icve/blob/master/update.html").text
        soup = BeautifulSoup(html, 'lxml')
        new_version = soup.find(class_="line", id="LC1").text
        if version != new_version[:6]:
            print("本程序有新版本，最新版本为%s，请前往GitHub或者Gitee下载最新版本" % new_version[:6])
            content = input("继续使用旧版本请输入“否”\n")
            if content == "否":
                version_choise()
            else:
                exit()
        else:
            print("已是最新版本")
            version_choise()


def version_choise():
    version_choice = input("“旧版MOOC”扣1，“新版MOOC”扣2\n")
    if version_choice == "1":
        old = OldMOOC()
        old.main()
    elif version_choice == "2":
        new = NewMOOC()
        new.main()
    else:
        version_choise()


if __name__ == "__main__":
    welcome()
