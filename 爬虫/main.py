from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook


class get_web:
    def __init__(self, url='https://www.nm.zsks.cn/'):

        self.url = url
        self.Info = {}
        # 规避被检测的风险
        option = ChromeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        broswer = webdriver.Chrome(executable_path="C:\\Users\\周琪雨\\Desktop\\爬虫\\chromedriver.exe", options=option)
        self.broswer = broswer
        self.wait = WebDriverWait(broswer, 5)

    def clk(self, btn):
        self.broswer.execute_script("arguments[0].click();", btn)

    def find_xpath(self, s):
        return self.wait.until(
            EC.visibility_of_element_located((By.XPATH, s))
        )

    def run(self):

        work = Workbook()
        sheet = work.active
        tittle = ['考生号', '性名', '性别', '民族', '考生状态', '院校名称', '总分', '加分条件', '特征分', '录取层次', '录取院校',
                  '录取专业', '录取方式', '录取时间', '报考层次', '填报次序', '报考院校', '报考专业1', '报考专业2', '报考专业3',
                  '报考专业4', '报考专业5', '报考专业6', '服从调剂', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩',
                  '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩',
                  '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩']
        sheet.append(tittle)
        work.save(filename="data.xlsx")


        self.broswer.get(self.url)
        ptgk = self.find_xpath('/html/body/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[2]/td/a[2]')
        self.clk(ptgk)

        ptgk2 = self.find_xpath('/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table[12]/tbody/tr/td/table[2]/tbody/tr/td[1]/a[7]')
        self.clk(ptgk2)

        # sleep(0.2)
        windows = self.broswer.window_handles
        self.broswer.switch_to.window(windows[-1])

        s1 = 2
        while s1 <= 100 :

            try:
                select_1 = self.find_xpath('/html/body/center/font[1]/font/form[1]/select/option[{}]'.format(s1))
                select_1.click()
                # sleep(0.2)

                s2 = 2
                while s2 <= 100:

                    try:
                        select_2 = self.find_xpath('/html/body/center/font[1]/font/form[2]/select/option[{}]'.format(s2))
                        select_2.click()
                        # sleep(0.2)
                        select_3 = self.find_xpath('/html/body/center/font[1]/font/form[3]/select/option[3]')
                        select_3.click()
                        # sleep(0.2)

                        s4 = 1
                        while s4 <= 100:

                            try:
                                select_4 = self.find_xpath('/html/body/center/font[1]/font/form[4]/select/option[{}]'.format(s4))
                                select_4.click()
                                # sleep(0.2)
                                self.find_xpath('/html/body/center/font[1]/font/form[4]/input[1]').click()

                                # sleep(0.2)
                                zy_tab = self.find_xpath('/html/body/center/p[2]/table')#获取专业表格
                                row_zy_tab = len(zy_tab.find_elements_by_tag_name('tr'))#获取表格有多少行 首先是获取表格的所有的tr标记的行返回值{}{}{}然后要把这个转化成对应的数字表示长度
                                #枚举表格的行
                                for zy in range(2,row_zy_tab+1):
                                    #获取表格对应的位置的连接
                                    zy_name = self.find_xpath('/html/body/center/p[2]/table/tbody/tr[{}]/td[2]/p/a'.format(zy))
                                    zy_name.click()
                                    # sleep(0.2)
                                    #获取表格的行
                                    number = self.find_xpath('/html/body/center/p/table')
                                    row_number = len(number.find_elements_by_tag_name('tr'))
                                    #进行循环
                                    for num in range(2,row_number+1):
                                        #获取对应的链接
                                        number = self.find_xpath('/html/body/center/p/table/tbody/tr[{}]/td[1]/p'.format(num))
                                        number.click()
                                        # sleep(0.2)
                                        #用一个list来存数据
                                        studentInfo = []
                                        #循环对应的3个表格，因为前3个表格是固定的，并且tr代表行，td代表列，只需要枚举对应的列就行
                                        for item in range(1, 10):
                                            studentInfo.append(self.find_xpath('/html/body/center/table[1]/tbody/tr[2]/td[{}]/p'.format(item)).text)

                                        for item in range(1,6):
                                            studentInfo.append(self.find_xpath('/html/body/center/table[2]/tbody/tr[2]/td[{}]/p'.format(item)).text)

                                        for item in range(1, 11):
                                            studentInfo.append(self.find_xpath('/html/body/center/table[3]/tbody/tr[2]/td[{}]/p'.format(item)).text)
                                        #获取对应的表格
                                        row = self.find_xpath('/html/body/center/table[4]')
                                        #item_i获取对应表格的总行，item_j获取对应表格的总列
                                        item_i = len(row.find_elements_by_tag_name('tr'))
                                        item_j = len(row.find_elements_by_tag_name('td'))
                                        local = 12 #因为每一行一共有12列，我们需要从第二行获取数据，所以从第二行开始枚举每次用总共的减去前一列的元素总和
                                        #枚举行
                                        for i in range (2,item_i+1):
                                            item_j -= local
                                            if item_j <= 0:
                                                break
                                            for j in range(1,min(item_j+1,local+1)):
                                                #把对应的数据加入你统计数据的数组
                                                studentInfo.append(self.find_xpath('/html/body/center/table[4]/tbody/tr[{}]/td[{}]/p'.format(i, j)).text)

                                        # print(studentInfo)
                                        #sheet是excel的标签，把对应的数据加入excel文档中
                                        sheet.append(studentInfo)
                                        # sleep(0.2)
                                        #保存文件
                                        work.save(filename="data.xlsx")
                                        #找到对应的返回按钮点击
                                        goback = self.find_xpath('/html/body/center/input')
                                        goback.click()
                                        # sleep(0.2)
                                    #找到对应的返回按钮点击
                                    goback = self.find_xpath('/html/body/center/p/input')
                                    goback.click()
                                    # sleep(0.2)

                            except TimeoutException:
                                break

                            if s4 == 1:
                                s4 = 3
                            else:
                                s4 += 1

                    except TimeoutException:
                        break

                    if s2 == 2:
                        s2 = 4
                    else:
                        s2 += 1

            except TimeoutException:
                break

            if s1 == 2:
                s1 = 4
            else:
                s1 += 1


if __name__ == '__main__':
    web = get_web()
    web.run()