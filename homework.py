from time import sleep
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
import xlwt
import xlrd
from xlutils.copy import copy

# driver加载浏览器驱动
driver = Chrome()
# filepath指示输出的文件路径
filepath = r"C:\Users\Administrator\Desktop\homework.xls"
# cnt表示当前excel表格中的sheet数量
cnt = 0
# point_num表示限定的sheet数量
point_num = 50
# source指示操作的sheet是否为源
source = True
# flag指示已经经过处理的sheet数量
flag = 0
# signal指示excel表中的sheet数量是否已经达到最大
signal = 0
# style定义表格样式
style0 = xlwt.XFStyle()
style1 = xlwt.XFStyle()
# dealed_list存放已经处理过的sheet
dealed_list = []


# 判断目标元素是否存在
def element_exist(finder, method, para):
    result = []
    if method == "ID":
        result = finder.find_elements(By.ID, para)
    if method == "CSS_SELECTOR":
        result = finder.find_elements(By.CSS_SELECTOR, para)
    if method == "CLASS_NAME":
        result = finder.find_elements(By.CLASS_NAME, para)
    if method == "XPATH":
        result = finder.find_elements(By.XPATH, para)
    if len(result) > 0:
        return True
    else:
        return False


# 爬取稿件数据并输出到文件
def get_video(url):
    new_workbook = copy(xlrd.open_workbook(filepath, formatting_info=True))
    sh = new_workbook.get_sheet(cnt)
    driver.get(url+"video")
    sleep(1)
    print("投稿情况")
    if not element_exist(driver, "CSS_SELECTOR", "[class='small-item fakeDanmu-item']"):
        print("投稿为空")
        sh.write_merge(1, 1, 0, 1, "投稿", style1)
    else:
        video_elements = driver.find_elements(By.CSS_SELECTOR, "[class='small-item fakeDanmu-item']")
        video_list = []
        href_list = []
        for element in video_elements:
            if element_exist(element, "XPATH", ".//a[2]"):
                ele = element.find_element(By.XPATH, ".//a[2]")
                video = ele.text
                href = ele.get_attribute("href")
                print(video, end=" ")
                print(href)
                video_list.append(video)
                href_list.append(href)
            else:
                print("找不到投稿")
                video_list.append("")
                href_list.append("")
        sh.write_merge(1, 1, 0, 1, "投稿", style1)
        for row in range(len(video_list)):
            sh.write(row+2, 0, video_list[row])
        for row in range(len(href_list)):
            sh.write(row+2, 1, href_list[row])
    new_workbook.save(filepath)


# 爬取频道数据并输出到文件
def get_channel(url):
    new_workbook = copy(xlrd.open_workbook(filepath, formatting_info=True))
    sh = new_workbook.get_sheet(cnt)
    driver.get(url+"channel/index")
    sleep(1)
    print("频道情况")
    if not element_exist(driver, "CLASS_NAME", "channel-item"):
        print("频道为空")
        sh.write_merge(1, 1, 2, 3, "频道", style1)
    else:
        channel_elements = driver.find_elements(By.CLASS_NAME, "channel-item")
        channel_list = []
        href_list = []
        for element in channel_elements:
            if element_exist(element, "CLASS_NAME", "channel-name"):
                ele = element.find_element(By.CLASS_NAME, "channel-name")
                channel = ele.text
                print(channel, end=" ")
                channel_list.append(channel)
            else:
                print("找不到频道")
                channel_list.append("")
            if element_exist(element, "XPATH", ".//a"):
                ele = element.find_element(By.XPATH, ".//a")
                href = ele.get_attribute("href")
                print(href)
                href_list.append(href)
            else:
                print("找不到链接")
                href_list.append("")
        sh.write_merge(1, 1, 2, 3, "频道", style1)
        for row in range(len(channel_list)):
            sh.write(row+2, 2, channel_list[row])
        for row in range(len(href_list)):
            sh.write(row+2, 3, href_list[row])
    new_workbook.save(filepath)


# 爬取关注数据并输出到文件
def get_follow(url):
    new_workbook = copy(xlrd.open_workbook(filepath, formatting_info=True))
    sh = new_workbook.get_sheet(cnt)
    driver.get(url+"fans/follow")
    sleep(1)
    print("关注情况")
    if not element_exist(driver, "CSS_SELECTOR", "[class='list-item clearfix']"):
        print("关注为空")
        sh.write_merge(1, 1, 4, 6, "关注", style1)
    else:
        follow_elements = driver.find_elements(By.CSS_SELECTOR, "[class='list-item clearfix']")
        name_list = []
        action_list = []
        href_list = []
        for element in follow_elements:
            if element_exist(element, "XPATH", ".//div[2]/a"):
                ele = element.find_element(By.XPATH, ".//div[2]/a")
                name = ele.text
                href = ele.get_attribute("href")
                print(name, end=" ")
                print(href, end=" ")
                name_list.append(name)
                href_list.append(href)
            else:
                print("找不到关注")
                name_list.append("")
                href_list.append("")
            if element_exist(element, "XPATH", ".//div[2]/p"):
                ele = element.find_element(By.XPATH, ".//div[2]/p")
                action = ele.text
                print(action)
                action_list.append(action)
            else:
                print("找不到描述")
                action_list.append("")
        sh.write_merge(1, 1, 4, 6, "关注", style1)
        for row in range(len(name_list)):
            sh.write(row+2, 4, name_list[row])
        for row in range(len(action_list)):
            sh.write(row+2, 5, action_list[row])
        for row in range(len(href_list)):
            sh.write(row+2, 6, href_list[row])
    new_workbook.save(filepath)


def initial_style():
    global style0
    global style1
    style0 = xlwt.easyxf("font: height 400; align: vert centre, horiz centre")
    style1 = xlwt.easyxf("font: height 300; align: vert centre, horiz centre")


# 设置sheet单元格大小
def set_sheet(sheet):
    sheet.col(0).width = 256 * 60
    sheet.col(1).width = 256 * 45
    sheet.col(2).width = 256 * 30
    sheet.col(3).width = 256 * 60
    sheet.col(4).width = 256 * 25
    sheet.col(5).width = 256 * 40
    sheet.col(6).width = 256 * 35


# 根据url增加一个sheet
def deal_url(father, url, count):
    global cnt
    global point_num
    global source
    global signal
    if cnt > point_num:
        signal = 1
        return
    else:
        if source:
            source = False
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet("浙江大学")
            set_sheet(sheet)
            sheet.write_merge(0, 0, 0, 6, "浙江大学", style0)
            workbook.save(filepath)
            get_video(url)
            get_channel(url)
            get_follow(url)
        else:
            driver.get(url)
            sleep(1)
            if not element_exist(driver, "ID", "h-name"):
                print("找不到主页名称")
                host_name = "找不到的主页名称"
            else:
                host_element = driver.find_element(By.ID, "h-name")
                host_name = host_element.text
                print(host_name)
            workbook = copy(xlrd.open_workbook(filepath, formatting_info=True))
            sheet_name = father+" 关注"+str(count)+" "+host_name
            sheet = workbook.add_sheet(sheet_name)
            set_sheet(sheet)
            sheet.write_merge(0, 0, 0, 6, host_name, style0)
            cnt += 1
            workbook.save(filepath)
            get_video(url)
            get_channel(url)
            get_follow(url)


# 根据一个sheet增加其中关注url的对应sheets
def deal_sheet(index):
    global signal
    global dealed_list
    data = xlrd.open_workbook(filepath)
    table = data.sheet_by_index(index)
    father = table.cell_value(0, 0)
    if father in dealed_list:
        return
    else:
        dealed_list.append(father)
        urls = table.col_values(6, 2, None)
        if urls is not None:
            for count in range(len(urls)):
                if urls[count] != "" and signal == 0:
                    deal_url(father, urls[count], count+1)
                else:
                    return


def run(x):
    global flag
    global signal
    initial_style()
    deal_url("", "https://space.bilibili.com/378951562/", 0)
    if signal == 1:
        print("sheet已达最大数量！")
        return
    while x > 0:
        sheets_num = xlrd.open_workbook(filepath).nsheets
        for flag in range(flag, sheets_num):
            deal_sheet(flag)
            if signal == 1:
                print("sheet已达最大数量！")
                return
        flag += 1
        x -= 1


# 运行程序
run(2)
