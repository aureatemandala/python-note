#-*- coding=utf-8 -*-


from openpyxl import load_workbook          
from openpyxl import Workbook  
from bs4 import BeautifulSoup
import re
import openpyxl
import requests
from selenium import webdriver
import json
import demjson
import os
from lxml import etree


def get_excel(bookname_list):
    "将excel表中的数据加载至列表"

    
    workbook = load_workbook(filename='./douban/files/1.xlsx')
    sheet = workbook['Sheet1']
    for cell in sheet['B']:
        bookname_list.append(cell.value)
    return bookname_list



def get_subject_id(wd, bookname):
    "搜索关键词，获取subject_id"


    url = 'https://search.douban.com/book/subject_search?search_text=' + bookname
    wd.get(url)

    element = wd.find_element_by_css_selector('.cover-link')
    datalist =  element.get_attribute('data-moreurl')
    datalist = datalist[22:-2]
    subjeci_id = demjson.decode(datalist)['subject_id']

    return subjeci_id


def getdata(subject_id,book_data_list):
    "此函数用来接收并处理HTML文件"

    #定义书本字典
    book_data = {
        "ISBN" : "",
        "作者" : "",
        "出版社" : "",
        "出品方" : "",
        "副标题" : "",
        "原作名" : "",
        "译者" : "",
        "出版年" : "",
        "页数" : "",
        "定价" : "",
        "装帧" : "",
        "丛书" : "",
        "国籍" : "",
        "简介" : "",
        "封面" : ""
    }

    #接收HTML文件
    diy_headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
    }

    url = "https://book.douban.com/subject/" + subject_id
    result = requests.get(url,headers = diy_headers)
    soup_html = result.text
    soup = BeautifulSoup(soup_html, 'html.parser')

    #处理图片链接
    img_html = str(soup.find_all(class_='nbg'))
    img_link_form = re.compile(r"https.*?l/public/.*?.jpg")
    img_link = re.findall(img_link_form,img_html)[0]
    book_data['封面'] = img_link
    
    #处理其他信息
    info_html = str(soup.find_all(id="info"))
    pass


    return



def test():


    #定义书本字典
    book_data = {
        "ISBN" : "",
        "作者" : "",
        "出版社" : "",
        "出品方" : "",
        "副标题" : "",
        "原作名" : "",
        "译者" : "",
        "出版年" : "",
        "页数" : "",
        "定价" : "",
        "装帧" : "",
        "丛书" : "",
        "国籍" : "",
        "简介" : "",
        "封面" : ""
    }


    diy_headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
    }
    url = "https://book.douban.com/subject/6082808/"
    result = requests.get(url,headers = diy_headers)
    soup_html = result.text
    soup = BeautifulSoup(soup_html, 'html.parser')
    info_html = str(soup.find_all(id="info"))
    

    #处理简单信息
    selector=etree.HTML(info_html)
    datasoup = selector.xpath('//span')
    for item in datasoup:
        if item.text == '出版社:':
            book_data['出版社'] = item.tail.lstrip()
        elif item.text == '原作名:':
            book_data['原作名'] = item.tail.lstrip()
        elif item.text == '出版年:':
            book_data['出版年'] = item.tail.lstrip()
        elif item.text == '页数:':
            book_data['页数'] = item.tail.lstrip()
        elif item.text == '定价:':
            book_data['定价'] = item.tail.lstrip()
        elif item.text == '装帧:':
            book_data['装帧'] = item.tail.lstrip()
        elif item.text == 'ISBN:':
            book_data['ISBN'] = item.tail.lstrip()
        elif item.text == '副标题:':
            book_data['副标题'] = item.tail.lstrip()
        else:
            pass
    
    #处理作者
    author_form = re.compile("作者.*?<.*?br.*?>",flags=re.S)
    author_list = str(re.findall(author_form,info_html))
    selector2=etree.HTML(author_list)
    datasoup2 = selector2.xpath('//a')
    for item in datasoup2:
        print(item.text)

    # ret = selector.xpath('//a')
    # for item in ret:
    #     print(item.text)


    return



def changedata(datalist):
    pass

def savedata():
    pass

def saveimg(url):
    "保存图片"

    diy_headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36',
    }
    picture = requests.get(url,headers = diy_headers,stream = True)
    print(picture.status_code) # 返回状态码
    if picture.status_code == 200:
        open('img.jpg', 'wb').write(picture.content) # 将内容写入图片
        print("done")
    del picture
    return


def main():
    #建立一个包含书名的列表bookname_list
    bookname_list = []
    bookname_list = get_excel(bookname_list)

    book_data_list = []

    
    wd = webdriver.Chrome("./douban/chromedriver.exe")
    for bookname in bookname_list:
        subject_id =  get_subject_id(wd, bookname)
        getdata(subject_id, book_data_list)

    wd.quit()


if __name__ == "__main__":
    #main()
    test()