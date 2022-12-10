from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import xlrd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys


def xls(name):
    workbook = xlrd.open_workbook("**********.xlsx")
    worksheet = workbook.sheet_by_name("角色")
    rows = worksheet.nrows
    for i in range(rows):
        if (name == worksheet.cell_value(i, 1)):
            return ([
                worksheet.cell_value(i, 5),
                worksheet.cell_value(i, 7),
                worksheet.cell_value(i, 8),
                worksheet.cell_value(i, 9),
                worksheet.cell_value(i, 10),
                worksheet.cell_value(i, 11),
                worksheet.cell_value(i, 12),
                worksheet.cell_value(i, 13),
                worksheet.cell_value(i, 14)
            ])
    return (False)


def Loop():
    time.sleep(2)
    windows = driver.window_handles
    driver.switch_to.window(windows[-1])
    items = driver.find_elements(By.CLASS_NAME, "box-unit")
    time.sleep(2)
    for i in range(len(items)):
        items[i].click()
        dialog = driver.find_element(By.CLASS_NAME, "rz-dialog-wrapper")
        list = xls(dialog.find_element(By.CLASS_NAME, "rz-dialog-title").text)
        print(list)
        if (list):
            time.sleep(1)
            dialog.find_element(By.NAME, "Level").send_keys(Keys.CONTROL + 'a')
            time.sleep(1)
            dialog.find_element(By.NAME, "Level").send_keys(list[0])
            dialog.find_element(By.NAME,
                                "Rarity").send_keys(Keys.CONTROL + 'a')
            time.sleep(1)
            dialog.find_element(By.NAME, "Rarity").send_keys(list[1])
            dialog.find_element(By.NAME,
                                "Promotion").send_keys(Keys.CONTROL + 'a')
            time.sleep(1)
            dialog.find_element(By.NAME, "Promotion").send_keys(list[2])
            equipment = dialog.find_elements(By.CLASS_NAME, "unit-equip-img")
            time.sleep(1)
            dialog.find_element(By.CLASS_NAME, "rz-dialog-title").click()
            time.sleep(1)
            for k in range(len(equipment)):
                print(list[3 + k])
                if (list[3 + k]):
                    equipment[k].click()
                    dialog.find_element(By.CLASS_NAME,
                                        "rz-dialog-title").click()
                    time.sleep(0.5)
        save = dialog.find_elements(By.CLASS_NAME, "rz-button-box")
        save[3].click()


option = webdriver.ChromeOptions()
option.add_experimental_option("detach", True)
driver = webdriver.Chrome('D:\driver\chromedriver.exe', options=option)

driver.get('https://pcr.satroki.tech/')
time.sleep(10)
driver.find_element(By.NAME, "UserName").send_keys('*******')
driver.find_element(By.NAME, "Password").send_keys('*********')
driver.find_element(By.CLASS_NAME, "rz-button-text").click()

Loop()
