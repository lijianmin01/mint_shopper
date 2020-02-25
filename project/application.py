from tkinter import *
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from pyquery import PyQuery as pq
import xlsxwriter

# 用chrome浏览器
browser = webdriver.Chrome()
wait = WebDriverWait(browser,10)

# 表格位置记录
# 列号
NUM=0 # 条数
COL = 0


def get_first_page(commodity):
    try:
        url="https://www.jd.com"
        browser.get(url)
        input=browser.find_element_by_id("key")
        input.send_keys(commodity)
        cnt=browser.find_element_by_css_selector("#search > div > div.form > button")
        cnt.click()
        #获取页数
        browser.implicitly_wait(3)
        pages=browser.find_element_by_xpath('//*[@id="J_bottomPage"]/span[2]/em[1]/b')
        print(pages)
        return int(pages.text)
    except TimeoutError:
        return get_first_page(commodity)

def deal_with_html():
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,'#J_goodsList > ul ')))
    html=browser.page_source
    doc=pq(html)
    items=doc('#J_goodsList > ul > li').items()
    for item in items:
        global ROW
        ROW += 1
        product={
            'href':item.find('div > div.p-img > a').attr('href'),
            'price':item.find('.p-price').text(),
            'p-tag':item.find('.p-tag').text(),
            'skcolor_ljg':item.find('.skcolor_ljg').text(),
            'em':item.find('div > div.p-name.p-name-type-2 > a > em').text()
        }
        if product['href'][0] != 'h':
            product['href'] = 'https:' + product['href']


        worksheet.write(ROW, COL, ROW-1)
        worksheet.write(ROW, COL + 1, product['p-tag'])
        worksheet.write(ROW, COL + 2, product['skcolor_ljg'])
        worksheet.write(ROW, COL + 3, product['em'])
        worksheet.write(ROW, COL + 4, product['price'])
        worksheet.write(ROW, COL + 5, product['em'])
        worksheet.write(ROW, COL + 6, product['href'])


def change_page(nums):
    time.sleep(3)
    input_page_num=browser.find_element_by_xpath('//*[@id="J_bottomPage"]/span[2]/input')
    input_page_num.send_keys(u'\ue003')
    input_page_num.send_keys(nums)
    cnt_click=browser.find_element_by_xpath('//*[@id="J_bottomPage"]/span[2]/a')
    cnt_click.click()
    global show
    show['text']="正在处理"+str(nums)+"网页"
    deal_with_html()

def main(commodity):
    try:
        global ROW
        ROW=0
        get_first_page(commodity)
        # 实现页面的跳转
        global total
        total=total.get()
        for i in range(int(total)):
            change_page(i+2)

        workbook.close()
        global show
        browser.close()
        show['text']="搜索完成"
    except Exception as e:
        show['text']=str(e)
        print('原因：'+str(e))
        browser.close()

def create_sheet(commodity_name):
    global workbook
    global worksheet
    workbook = xlsxwriter.Workbook(commodity_name)
    worksheet = workbook.add_worksheet()
    worksheet.write(0,COL,'')
    worksheet.write(0, COL + 1, '')
    worksheet.write(0, COL + 2, '产品类型')
    worksheet.write(0, COL + 3, '分类')
    worksheet.write(0, COL + 4, '价格')
    worksheet.write(0, COL + 5, '信息')
    worksheet.write(0, COL + 6, '链接')



class Application(Frame):

    def __init__(self,master=None):
        super().__init__(master)
        self.master=master
        self.pack()
        self.createWidget()

    def createWidget(self):
        Label(self,text="~~~~~~~~~~").grid(row=0)
        global commdity,total
        Label(self, text="商品名称").grid(row=1)
        Label(self, text="删选数量").grid(row=2)

        e1 = Entry(self)
        e2 = Entry(self)
        commdity = e1
        total = e2
        e1.grid(row=1, column=1)
        e2.grid(row=2, column=1)

        button2 = Button(self, text='Start', command=self.start_app)
        button2.grid(row=3, column=0)

        button1 = Button(self, text='END',command=self.end_app)
        button1.grid(row=3, column=2)
        global show
        show = Label(root, width=40, height=3, bg="#fff")
        show.pack()
        show['text']=""
        mainloop()

    def start_app(self):
        global show
        show['text']="运行中"
        global commdity
        print(commdity.get())
        commodity=commdity.get()
        commodity_name=commodity+".xlsx"
        create_sheet(commodity_name)
        main(commodity)

    def end_app(self):
        global show
        show['text']='程序已停止执行'

    def suc_app(self):
        global show
        browser.close()
        show['text']="运行成功"



if __name__=='__main__':
    root=Tk()
    global ROW,show
    NUM = 0
    root.geometry("360x180")
    app=Application(master=root)
    root.mainloop()

# if __name__ == '__main__':
