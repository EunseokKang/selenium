from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert
import os
import time
import csv
import shutil
import csv

def extract_column_values(csv_file, column_name):
    values = []
    
    with open(csv_file, 'r', newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        
        for row in reader:
            value = row[column_name]
            
            # 빈 값인 경우 제외
            if value.strip():  # 문자열의 앞뒤 공백을 제거하고 난 후에 비어있지 않은 경우
                values.append(value)
    
    return values

# CSV 파일 경로와 추출할 열(column) 이름 지정
csv_file_path = 'file.csv'
column_name = '정류장명'

# 정류장명 열의 요소들을 리스트로 추출 (빈 값 제외)
bus_stop_name = extract_column_values(csv_file_path, column_name)


url = "https://www.stcis.go.kr/pivotIndi/wpsPivotIndicator.do?siteGb=P&indiClss=IC03&indiSel=IC0304"
dr = webdriver.Chrome()
dr.get(url)
time.sleep(5)

# 기간 선택
def select_month():
    wait = WebDriverWait(dr, 1000)
    element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "searchDateGubun")))
    day_select = dr.find_element(By.CLASS_NAME, "searchDateGubun")
    day_select.click()
    time.sleep(1)

    calander = dr.find_element(By.ID, "searchFromMonth")
    calander.click()
    time.sleep(1)

    month_select = dr.find_element(By.CLASS_NAME, "ui-datepicker-month")
    month_select.click()
    time.sleep(1)

    month_sel = Select(month_select)
    month_sel.select_by_value("0")
    time.sleep(1)

    confirm = dr.find_element(By.CLASS_NAME, "ui-datepicker-close.ui-state-default.ui-priority-primary.ui-corner-all")
    confirm.click()
    time.sleep(1)

# 구 찾기
def search_csv(input_value):
    result = []
    with open('file.csv', newline='', encoding='cp949') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row['정류장명'] == input_value:
                result = row['시군구명']
                break
    return result

# 공간 선택
def space_select(abc):
    wait = WebDriverWait(dr, 1000)
    search_box = dr.find_element(By.NAME, "searchSttnSpaceNm")
    search_box.click()
    time.sleep(1)

    global a
    a = search_csv(abc)
    search_box.send_keys(a)
    time.sleep(2)
    actions = ActionChains(dr)
    actions.send_keys(Keys.ENTER).perform()
    element = wait.until(EC.presence_of_element_located((By.ID, "divSttn")))

    if a == '서구':
        gu = dr.find_element(By.ID, "label_30170")
        gu.click()
    elif a == '중구':
        gu = dr.find_element(By.ID, "label_30140")
        gu.click()
    elif a == '동구':
        gu = dr.find_element(By.ID, "label_30110")
        gu.click()
    elif a == '유성구':
        gu = dr.find_element(By.ID, "label_30200")
        gu.click()
    elif a == '대덕구':
        gu = dr.find_element(By.ID, "label_30230")
        gu.click()
    time.sleep(2)

#정류장 조회 결과
def stop_inquiry_result(abc):
    search_box = dr.find_element(By.NAME, "popupSearchSttnNma")
    search_box.click()
    time.sleep(1)
    search_box.send_keys(abc)
    time.sleep(1)

    actions = ActionChains(dr)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(1)

#정류장 조회 결과 2
def stop_inquiry_result_2(abc):
    wait = WebDriverWait(dr, 1000)
    element = wait.until(EC.presence_of_element_located((By.ID, "chkSttn0")))

    outer = dr.find_elements(By.CLASS_NAME, "check_area.pos")
    a = len(outer)
    b = list(range(a))
    count = 1
    for i in b:
        if i > 0:
            select_month()
            space_select(abc)
            stop_inquiry_result(abc)
            element = wait.until(EC.presence_of_element_located((By.ID, "chkSttn0")))

        outer = dr.find_elements(By.CLASS_NAME, "check_area.pos")

        if outer[i+1].find_element(By.XPATH, """//*[@id="chkSttn%s"]""" % (i)).get_attribute("title") == abc:
            button = outer[i+1].find_element(By.CLASS_NAME, "check")
            button.click()
            time.sleep(1)
            select_button = dr.find_elements(By.CLASS_NAME, "but_box.blue")
            select_button[2].click()
            time.sleep(1)
            select_button[0].click()
            time.sleep(1)
            da = Alert(dr)
            da.accept()
            time.sleep(1)

            #결과 다운로드
            element = wait.until(EC.presence_of_element_located((By.ID, "divWarning")))
            time.sleep(5)
            download_button = dr.find_elements(By.XPATH, """//*[@id="btnExport"]""")
            dr.execute_script("rgrstyExcelExport();")
            time.sleep(5)

            time.sleep(10)

            dir="csvFiles"
            dir2 = f"csvFiles/{abc}"
            if not os.path.exists(dir):
                os.makedirs(dir)
            if not os.path.exists(dir2):
                os.makedirs(dir2)

            downloads_path = "C:/Users/unseo/Downloads"
            dest_folder = f"D:/Coding/selenium/csvFiles/{abc}"
            files = os.listdir(downloads_path)
            files = [(os.path.join(downloads_path, f), os.stat(os.path.join(downloads_path, f)).st_mtime) for f in files]
            files.sort(key=lambda x: x[1])
            latest_file_path = files[-1][0]
            new_file_path = os.path.join(dest_folder, f"{abc}_{count}월.csv")
            shutil.copy2(latest_file_path, new_file_path)
            os.remove(latest_file_path)

            back = dr.find_element(By.CLASS_NAME, "but_box")
            dr.execute_script("fnDisplayResult('close');")
            time.sleep(1)
            count+=1

            #뒤로가기 후 월 재선택
            for i in range(5):
                calendar = dr.find_element(By.ID, "searchFromMonth")
                calendar.click()
                element5 = wait.until(EC.presence_of_element_located((By.XPATH, """//*[@id="ui-datepicker-div"]/div[1]/a[2]""")))
                time.sleep(5)

                month_select = dr.find_element(By.XPATH, """//*[@id="ui-datepicker-div"]/div[1]/div/select[2]""")
                month_select.click()
                time.sleep(1)

                month_sel = Select(month_select)
                month_sel.select_by_value(str(i+1))
                time.sleep(1)

                confirm = dr.find_element(By.CLASS_NAME, "ui-datepicker-close.ui-state-default.ui-priority-primary.ui-corner-all")
                confirm.click()
                time.sleep(1)

                search_button = dr.find_element(By.XPATH, """//*[@id="btnSearch"]/button""")
                search_button.click()
                time.sleep(1)
                da.accept()
                time.sleep(1)
                element123 = wait.until(EC.presence_of_element_located((By.ID, "divWarning")))

                dr.execute_script("rgrstyExcelExport();")
                time.sleep(5)

                dir="csvFiles"
                dir2 = f"csvFiles/{abc}"
                if not os.path.exists(dir):
                    os.makedirs(dir)
                if not os.path.exists(dir2):
                    os.makedirs(dir2)

                downloads_path = "C:/Users/unseo/Downloads"
                dest_folder = f"D:/Coding/selenium/csvFiles/{abc}"
                files = os.listdir(downloads_path)
                files = [(os.path.join(downloads_path, f), os.stat(os.path.join(downloads_path, f)).st_mtime) for f in files]
                files.sort(key=lambda x: x[1])
                latest_file_path = files[-1][0]
                new_file_path = os.path.join(dest_folder, f"{abc}_{count}월.csv")
                shutil.copy2(latest_file_path, new_file_path)
                os.remove(latest_file_path)
                count += 1

                back2 = dr.find_element(By.CLASS_NAME, "but_box")
                dr.execute_script("fnDisplayResult('close');")
                element12 = wait.until(EC.presence_of_element_located((By.ID, "searchFromMonth")))
            dr.refresh()
            time.sleep(5)
        else:
            dr.refresh()
            time.sleep(5)
            continue


for j in bus_stop_name:
    select_month()
    space_select(j)
    stop_inquiry_result(j)
    try:
        stop_inquiry_result_2(j)
    except IndexError:
        dr.refresh()
        continue
    time.sleep(1)
    # dr.refresh()
# time.sleep(1000)