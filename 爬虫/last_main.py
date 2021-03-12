from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook


class get_web:
    def __init__(self, url='https://www.nm.zsks.cn/'):

        self.url = url
        self.Info = {}
        # 规避被检测的风险
        option = ChromeOptions()
        option.add_experimental_option('excludeSwitches', ['enable-automation'])
        broswer = webdriver.Chrome(executable_path="D:\\work\\爬虫\\chromedriver.exe", options=option)
        self.broswer = broswer
        self.wait = WebDriverWait(broswer, 5)

    def clk(self, btn):
        self.broswer.execute_script("arguments[0].click();", btn)

    def find_xpath(self, s):
        return self.wait.until(
            EC.visibility_of_element_located((By.XPATH, s))
        )

    def run(self):

        cont = 0
        work = Workbook()
        sheet = work.active
        tittle = ['考生号', '性名', '性别', '民族', '考生状态', '院校名称', '总分', '加分条件', '特征分', '录取层次', '录取院校',
                  '录取专业', '录取方式', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩','成绩代码', '成绩项', '成绩',
                  '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项',
                  '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码', '成绩项', '成绩', '成绩代码',
                  '成绩项', '成绩', '成绩代码', '成绩项', '成绩']
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
        while s1 <= 100:

            try:
                select_1 = self.find_xpath('/html/body/center/font[1]/font/form[1]/select/option[{}]'.format(s1))
                select_1.click()
                # sleep(0.05)

                s2 = 2
                while s2 <= 100:

                    try:
                        select_2 = self.find_xpath('/html/body/center/font[1]/font/form[2]/select/option[{}]'.format(s2))
                        select_2.click()
                        # sleep(0.05)
                        select_3 = self.find_xpath('/html/body/center/font[1]/font/form[3]/select/option[3]')
                        select_3.click()
                        # sleep(0.05)

                        s4 = 1
                        while s4 <= 1000:

                            try:
                                select_4 = self.find_xpath('/html/body/center/font[1]/font/form[4]/select/option[{}]'.format(s4))
                                select_4.click()
                                # sleep(0.05)
                                self.find_xpath('/html/body/center/font[1]/font/form[4]/input[1]').click()

                                # sleep(0.05)
                                zy_tab = self.find_xpath('/html/body/center/p[2]/table')
                                row_zy_tab = len(zy_tab.find_elements_by_tag_name('tr'))
                                for zy in range(2,row_zy_tab+1):

                                    try:

                                        zy_name = self.find_xpath('/html/body/center/p[2]/table/tbody/tr[{}]/td[2]/p/a'.format(zy))
                                        ActionChains(self.broswer).move_to_element(zy_name).click().perform()
                                        #让浏览器跳到对应的位置点击
                                        # zy_name.click()
                                        # sleep(0.05)

                                        number = self.find_xpath('/html/body/center/p/table')
                                        row_number = len(number.find_elements_by_tag_name('tr'))
                                        for num in range(2,row_number+1):

                                            try:

                                                number = self.find_xpath('/html/body/center/p/table/tbody/tr[{}]/td[1]/p'.format(num))
                                                ActionChains(self.broswer).move_to_element(number).click().perform()
                                                #让浏览器跳到对应的位置点击
                                                # number.click()
                                                # sleep(0.05)
                                                studentInfo = []

                                                for item in range(1, 10):
                                                    studentInfo.append(self.find_xpath('/html/body/center/table[1]/tbody/tr[2]/td[{}]/p'.format(item)).text)

                                                for item in range(1,5):
                                                    studentInfo.append(self.find_xpath('/html/body/center/table[2]/tbody/tr[2]/td[{}]/p'.format(item)).text)

                                                row = self.find_xpath('/html/body/center/table[4]')
                                                item_i = len(row.find_elements_by_tag_name('tr'))
                                                item_j = len(row.find_elements_by_tag_name('td'))
                                                local = 12
                                                for i in range (2,item_i+1):
                                                    item_j -= local
                                                    if item_j <= 0:
                                                        break
                                                    for j in range(1,min(item_j+1,local+1)):
                                                        studentInfo.append(self.find_xpath('/html/body/center/table[4]/tbody/tr[{}]/td[{}]/p'.format(i, j)).text)

                                                # print(studentInfo)
                                                sheet.append(studentInfo)
                                                cont += 1
                                                # sleep(0.05)
                                                print("complete " + str(cont))
                                                work.save(filename="data.xlsx")
                                                goback = self.find_xpath('/html/body/center/input')
                                                goback.click()
                                                sleep(0.05)

                                            except TimeoutException:
                                                continue

                                        goback = self.find_xpath('/html/body/center/p/input')
                                        goback.click()
                                        sleep(0.05)

                                    except TimeoutException:
                                        continue

                            except TimeoutException:
                                break

                            if s4 == 1:
                                s4 = 4
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