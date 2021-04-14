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
        broswer = webdriver.Chrome(executable_path="D:\\Desktop\\web-crawler\\爬取内蒙古高考教育网站\\chromedriver.exe", options=option)
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
        tittle = ['专业代号', '专业名称', '填报次序', '最高分', '最低分', '最低分位次', '录取人数','院校名称']
        sheet.append(tittle)
        work.save(filename="data.xlsx")

        self.broswer.get(self.url)
        ptgk = self.find_xpath('/html/body/table[2]/tbody/tr[1]/td[2]/table/tbody/tr[2]/td/a[2]')
        self.clk(ptgk)

        ptgk2 = self.find_xpath(
            '/html/body/table[4]/tbody/tr/td[2]/table[2]/tbody/tr/td[2]/table[12]/tbody/tr/td/table[2]/tbody/tr/td[1]/a[7]')
        self.clk(ptgk2)

        # sleep(0.2)
        windows = self.broswer.window_handles
        self.broswer.switch_to.window(windows[-1])

        s1 = 2
        while s1 <= 100:

            try:
                select_1 = self.find_xpath('/html/body/center/font[1]/font/form[1]/select/option[{}]'.format(s1))
                select_1.click()
                # sleep(0.2)

                s2 = 3
                while s2 <= 100:

                    try:
                        select_2 = self.find_xpath(
                            '/html/body/center/font[1]/font/form[2]/select/option[{}]'.format(s2))
                        select_2.click()
                        # sleep(0.2)
                        select_3 = self.find_xpath('/html/body/center/font[1]/font/form[3]/select/option[3]')
                        select_3.click()
                        # sleep(0.2)

                        s4 = 1
                        while s4 <= 100:

                            try:
                                select_4 = self.find_xpath(
                                    '/html/body/center/font[1]/font/form[4]/select/option[{}]'.format(s4))
                                select_4.click()
                                # sleep(0.2)
                                self.find_xpath('/html/body/center/font[1]/font/form[4]/input[1]').click()

                                # sleep(0.2)
                                zy_tab = self.find_xpath('/html/body/center/p[2]/table')  # 获取专业表格
                                row_zy_tab = len(zy_tab.find_elements_by_tag_name(
                                    'tr'))  # 获取表格有多少行 首先是获取表格的所有的tr标记的行返回值{}{}{}然后要把这个转化成对应的数字表示长度
                                # 枚举表格的行

                                for zy in range(2, row_zy_tab + 1):
                                    # 获取表格对应的位置的连接
                                    try:
                                        studentInfo = []
                                        for o in range(1, 9):
                                            if o == 2:
                                                try:
                                                    studentInfo.append(self.find_xpath(
                                                        '/html/body/center/p[2]/table/tbody/tr[{}]/td[{}]/p/a'.format(zy,
                                                                                                                      o)).text)
                                                except TimeoutException:
                                                    studentInfo.append(self.find_xpath(
                                                        '/html/body/center/p[2]/table/tbody/tr[{}]/td[{}]/p/a'.format(
                                                            zy - 1, o)).text)
                                            elif o == 6:
                                                try:
                                                    studentInfo.append(self.find_xpath(
                                                        '/html/body/center/p[2]/table/tbody/tr[{}]/td[{}]/p'.format(zy,
                                                                                                                    o)).text)
                                                except TimeoutException:
                                                    studentInfo.append("")
                                            elif o == 8:
                                                try:
                                                    studentInfo.append(self.find_xpath('/html/body/center/font[1]/font/form[4]/select/option[{}]'.format(s4)).text)
                                                    
                                                except TimeoutException:
                                                    studentInfo.append("")
                                            else:
                                                studentInfo.append(self.find_xpath(
                                                    '/html/body/center/p[2]/table/tbody/tr[{}]/td[{}]/p'.format(zy,
                                                                                                                o)).text)
                                            
                                    except TimeoutException:
                                        continue
                                    sheet.append(studentInfo)
                                    # sleep(0.2)
                                    # 保存文件
                                    work.save(filename="data.xlsx")


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