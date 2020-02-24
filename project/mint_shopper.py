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
            'href':'https:'+item.find('div > div.p-img > a').attr('href'),
            'price':item.find('.p-price').text(),
            'p-tag':item.find('.p-tag').text(),
            'skcolor_ljg':item.find('.skcolor_ljg').text(),
            'em':item.find('div > div.p-name.p-name-type-2 > a > em').text()
        }
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
    deal_with_html()

def main(commodity):
    global ROW
    ROW=0
    total=get_first_page(commodity)
    print(total)
    # 实现页面的跳转
    total=1
    for i in range(total):
        change_page(i+2)

    workbook.close()

def create_sheet(commodity_name):
    global workbook
    global worksheet
    workbook = xlsxwriter.Workbook(commodity_name)
    worksheet = workbook.add_worksheet()
    worksheet.write(0,COL,NUM)
    worksheet.write(0, COL + 1, '')
    worksheet.write(0, COL + 2, '产品类型')
    worksheet.write(0, COL + 3, '分类')
    worksheet.write(0, COL + 4, '价格')
    worksheet.write(0, COL + 5, '信息')
    worksheet.write(0, COL + 6, '链接')


if __name__ == '__main__':
    global ROW
    commodity='iphoneX'
    commodity_name=commodity+".xlsx"
    create_sheet(commodity_name)
    main(commodity)

'''
https://item.paipai.com/25549359209.html
'''