from urllib import  request
import sys
import io
from selenium import webdriver
from urllib import parse
from bs4 import BeautifulSoup
import time
import json
import xlwt
import requests
import os
from getpass import getpass

headers = {}  # 全局变量headers
student_name  = ""

def get_index(lists,strs):
    """
    该函数返回传经来的lists中的元素包含strs的元素索引
    """
    index = []
    for i in range(len(lists)):
        if strs in lists[i]:
            index.append(i)
    return index


def login():
    """
    实现用户的登陆
    """
    print('*' * 21, '欢迎使用太理教务查询系统', '*' * 21)
    print(' ' * 40, 'directed by Zhu Jiang')
    print('*' * 64)
    while True:
        student_id = input('输入学号:')
        password = getpass("输入密码(身份证后6位):")  # 该方法不会显示用户输入的密码，但是无法再IDE中的控制端运行，只能在cmd中实现
        print("正在登录。。。")
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')  # 改变标准输出的默认编码
        # 设置该options以使得selenium在运行是不会弹出浏览器框
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        browser = webdriver.Chrome(options=chrome_options)  # 初始化浏览器对象，options参数使得浏览器不会弹出
        base_url = "http://jxgl.tyut.edu.cn:999/"
        browser.get(base_url)
        browser.find_element_by_css_selector(".form-username ").send_keys(student_id)  # 帐号
        browser.find_element_by_css_selector(".form-password ").send_keys(password)  # 密码
        browser.find_element_by_css_selector(".btn").click()  # 登录
        time.sleep(2)  # 休眠2s来获取登陆后的页面，否则会返回登录页面的url
        current_url = browser.current_url
        if current_url != base_url:  # 如果当前url和base_url不一样，说明登录成功
            cookies = browser.get_cookies()  # 获取登陆后用户的cookies
            cookie_str = ""  # 构造cookie并实现拼接
            for item_cookie in cookies:
                item_str = item_cookie["name"]+"="+item_cookie["value"]+"; "
                cookie_str += item_str
            global headers
            # 将构造好的cookie放在headers并声明为全局变量以便后续的获取网页
            headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36",
                    "Cookie": cookie_str
            }
            # 获取网页信息并从中提取用户姓名
            demo = requests.get(current_url,headers=headers).text
            global student_name
            soup = BeautifulSoup(demo,'lxml')
            student_name = soup.select(".user-info")[0].get_text().split()[-1]
            print("登录成功,", student_name + " 同学，请输入你想获取什么信息？")
            browser.quit()
            return
        else:
            print("用户名或密码错误，请重新输入！")  # 输入错误后重新输入


def get_ranking():
    """
    获取用户的排名情况
    """
    try:
        url = "http://jxgl.tyut.edu.cn:999/Tschedule/C6Cjgl/GetXskccjResult"
        data = {'order': 'zxjxjhh desc,kch'}
        data = parse.urlencode(data).encode('gbk')
        req = request.Request(url, data=data, headers=headers)
        response = request.urlopen(req)
        html = response.read().decode('utf-8')
        soup = BeautifulSoup(html,'lxml')
        for i , j in zip(soup.select(".profile-info-name"),soup.select(".profile-info-value")):
            print(i.get_text(),end=" "+j.get_text())
            print()
    except Exception:
        print('获取成绩失败！')
        print("不好意思同学，大概率是学校网站崩了，或者是学校网站改动了，我会尽快处理的。")


def get_grades():
    """获取用户的成绩信息"""
    try:
        url = 'http://jxgl.tyut.edu.cn:999/Tschedule/C6Cjgl/GetKccjResult'
        data = {'order': 'zxjxjhh desc,kch'}
        data = parse.urlencode(data).encode('gbk')
        req = request.Request(url, data=data, headers=headers)
        response = request.urlopen(req)
        html = response.read().decode('utf-8')
        soup = BeautifulSoup(html, 'lxml')
        print("获取成绩成功！")
        print()
        courses = []
        for tr in soup.select("tr"):
            x = 0
            course = []
            for td in tr.select("td"):
                if x == 2 or x == 6:
                    course.append(td.get_text())
                elif '学年' in td.get_text():
                    course.append(td.get_text())
                x += 1
            courses.extend(course)
        index = get_index(courses,'学年')
        result_dict = {}
        for i in range(len(index)):
            if i+1!=len(index):
                result_dict.setdefault(courses[index[i]],[]).extend(courses[index[i]+1:index[i+1]])
            else:
                result_dict.setdefault(courses[index[i]], []).extend(courses[index[i]+1:])
        for i in result_dict:
            print("*"*20)
            print(i)
            print()
            turns = 0
            for j in result_dict[i]:
                if turns ==2:
                    print()
                    turns = 0
                print(format(j,"^25"),end="")
                turns += 1
            print()
            print()
    except Exception:
        print('获取排名失败！')
        print("不好意思同学，大概率是学校网站崩了，或者是学校网站改动了，我会尽快处理的。")


def get_course():
    filename = student_name + '同学的课程表.xlsx'
    if os.path.exists(filename):
        print("你已经生成课表了呦，请到对应目录查看！")
        return
    print('正在获取课表信息，即将生成课表文件。。。')
    url = 'http://jxgl.tyut.edu.cn:999/Tresources/A1Xskb/GetXsKb'
    data = {'zxjxjhh': ''}
    data = parse.urlencode(data).encode('gbk')
    req = request.Request(url, data=data, headers=headers)
    response = request.urlopen(req)
    html = response.read().decode('utf-8')
    soup = BeautifulSoup(html, 'lxml')
    course_dict = soup.p.string
    result = {}
    for i in range(1, 6):
        result[str(i)] = {}
    x = json.loads(course_dict)

    def course_detail(data):
        result = []
        result.append(data["Kcm"])
        result.append(data["Zcsm"])
        result.append(data["Dd"])
        result.append(data['Jsm'])
        result.append(data["Jc"])
        result.append(data['Skxq'])
        return result

    for i in x["rows"]:
        res = course_detail(i)
        if res[-1] != None:
            index = str(res[-1])
            result[index].setdefault(res[-2], []).append(res[:-2])
    date = ['星期一', '星期二', '星期三', '星期四', '星期五']
    course_time = ['1-2', '3-4', '5-6', '7-8']
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('sheet1')
    style = xlwt.XFStyle()  # 初始化样式
    style.alignment.wrap = 1  # 自动换行
    for i in range(8):
        first_col = sheet.col(i)  # xlwt中是行和列都是从0开始计算的
        first_col.width = 256 * 20
    tall_style = xlwt.easyxf('font:height 720;')
    for i in range(12):  # 36pt,类型小初的字号
        first_row = sheet.row(i)
        first_row.set_style(tall_style)

    for i in range(len(date)):
        sheet.write(0, i + 1, date[i], style)
    for i in range(len(course_time)):
        sheet.write(i + 1, 0, course_time[i], style)
    for i in result:
        for course_jie in result[i]:
            strs = ""
            for lengths in range(len(result[i][course_jie])):
                strs += "\n".join(result[i][course_jie][lengths])
                strs += "\n"
            sheet.write(int(int(course_jie[-1]) / 2), int(i), strs, style)
    book.save(filename)
    print("生成课表'"+filename+"'成功,请到对应目录查看！")



    
def main():
    login()
    while True:
        print()
        get_input = input("[1]获取排名/[2]获取课程成绩/[3]获取课表[0]退出:")
        if int(get_input) == 1:
            get_ranking()
        elif int(get_input) == 2:
            get_grades()
        elif int(get_input) == 3:
            get_course()
        elif int(get_input) == 0:
            print()
            print("退出成功，欢迎下次使用！")
            break
        else:
            print("同学，请输入1,2,3,0来选取对应内容！")


if __name__=="__main__":
    main()
