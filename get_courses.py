import time, os, json, random
import tqdm, openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager


def save_json_as_wb(data, filename, sheetname, main_grade=""):
    LABEL = [["修過", "學年度", "學期", "主開年級", "當期課號", "永久課號", "先修課程", "課程名稱", "學分", "時數", "老師", "課程時間/地點", "課程類別", "備註"],
            ["elected", "acy", "sem", "main_grade", "cos_id", "cos_code", "pre_req", "cos_cname", "cos_credit", "cos_hours", "teacher", "cos_time", "cos_type", "memo"],]
    FUNCTIONAL_LABLE = ["elected", "pre_req", "main_grade"]

    # 開啟工作表
    if not os.path.exists(filename):
        workbook = openpyxl.Workbook()
        sheet = workbook.worksheets[0]
        sheet.title = sheetname
        for i in LABEL:
            sheet.append(i)
    else:
        workbook = openpyxl.load_workbook(filename)
        if sheetname not in workbook.sheetnames:
            workbook.create_sheet(sheetname)
            sheet = workbook[sheetname]
            for i in LABEL:
                sheet.append(i)
        sheet = workbook[sheetname]
    
    # 記錄所有標籤以及他們的index之間的關係
    label_to_index = {}
    for i in range(1, sheet.max_column+1):
        label_to_index[sheet.cell(2, i).value] = i

    # 所有已經有的課程
    existed_courses = [sheet.cell(x, label_to_index["cos_id"]).value for x in range(3, sheet.max_row+1)]

    # 取得dep_id，e.g. "E2D47D2E-B529-4449-8619-8561D3401D32"
    dynamic_keys = list(data)
    for dynamic_key in tqdm.tqdm(dynamic_keys, "開課單位"):
        json_content = data[dynamic_key]

        course_dicts = [v for k, v in json_content.items() if k.isdigit()]
        for courses in course_dicts:
            for course in tqdm.tqdm(courses.values(), "儲存工作表", leave=False):
                if course.get('cos_id') not in existed_courses: # 避免重複的課程填入
                    course_info = []
                    for label in label_to_index:
                        if label in FUNCTIONAL_LABLE:
                            if label == "main_grade":
                                course_info.append(main_grade)
                            else:
                                course_info.append("")
                        else:
                            course_info.append(course.get(label))
                    sheet.append(course_info)
                    existed_courses.append(course.get('cos_id'))
        
    workbook.save(filename)

def fetch_from_timetable(reqs, filename, sheetname):
    options = Options()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    driver.get("https://timetable.nycu.edu.tw/")
    wait.until(lambda d: d.execute_script("return jQuery.active == 0"))

    hijack_script = """
    // 1. 準備一個全域變數來存結果
    window.captured_response = null;

    // 2. 備份原本的 XMLHttpRequest open 方法
    var originalOpen = XMLHttpRequest.prototype.open;

    // 3. 覆寫 open 方法
    XMLHttpRequest.prototype.open = function() {
        // 當請求完成載入時
        this.addEventListener('load', function() {
            // 過濾網址，只抓取我們想要的 API
            if (this.responseURL && this.responseURL.includes('get_cos_list')) {
                console.log("抓到了！: " + this.responseURL);
                window.captured_response = this.responseText; // 把結果存起來
            }
        });
        // 執行原本的 open
        originalOpen.apply(this, arguments);
    };
    """
    driver.execute_script(hijack_script)

    for req in reqs:
        try:
            print("執行要求:", req)
            
            main_year = ""
            for selection_id, selection_text in req:
                Select(driver.find_element(By.ID, selection_id)).select_by_visible_text(selection_text)
                if selection_id == 'fGrade':
                    main_year = selection_text
                wait.until(lambda d: d.execute_script("return jQuery.active == 0"))

            driver.execute_script("window.captured_response = null;")
            driver.find_element(By.ID, "crstime_search").click()

            # =================================================================
            # 等待並提取資料
            # =================================================================

            # 使用 wait.until 輪詢 JS 變數，直到 window.captured_response 有值
            # 這裡的意思是：每隔 0.5 秒問瀏覽器「window.captured_response 是不是 null？」
            raw_json = wait.until(lambda d: d.execute_script("return window.captured_response;"))
            
            if raw_json:        
                # 1. 解析 JSON 字串為 Python 字典
                data = json.loads(raw_json)
                save_json_as_wb(data, filename, sheetname, main_year)
            else:
                print("未抓取到資料")

            # 稍微等一下讓我們看結果
            if req != reqs[-1]:
                cd_time = random.randint(3, 10)
                for _ in tqdm.tqdm(range(cd_time), "等待一下"):
                    time.sleep(1)

        except Exception as e:
            print(f"發生錯誤: {e}")
    
    driver.quit()

fetch_from_timetable([[("fAcySem", "114 學年度 第 1 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "(醫學生物技術暨檢驗學系)"), ("fGrade", "全部")],
                      [("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "(醫學生物技術暨檢驗學系)"), ("fGrade", "全部")],],
                     "courses.xlsx", "MT")

fetch_from_timetable([[("fAcySem", "114 學年度 第 1 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "BME(生物醫學工程學系)"), ("fGrade", "全部")],
                      [("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "BME(生物醫學工程學系)"), ("fGrade", "全部")],],
                     "courses.xlsx", "BME")

fetch_from_timetable([[("fAcySem", "114 學年度 第 1 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "(數位醫療學士學位學程)"), ("fGrade", "全部")],
                      [("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生物醫學暨工程學院"), ("fDep", "(數位醫療學士學位學程)"), ("fGrade", "全部")],],
                     "courses.xlsx", "DHCR")

fetch_from_timetable([[("fAcySem", "114 學年度 第 1 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生命科學院"), ("fDep", "(生命科學系暨基因體科學研究所)"), ("fGrade", "全部")],
                      [("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "生命科學院"), ("fDep", "(生命科學系暨基因體科學研究所)"), ("fGrade", "全部")],],
                     "courses.xlsx", "LS")

fetch_from_timetable([[("fAcySem", "114 學年度 第 1 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "校級"), ("fDep", "(學士班大一大二不分系)"), ("fGrade", "全部")],
                      [("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班課程"), ("fCategory", "一般學士班"), ("fCollege", "校級"), ("fDep", "(學士班大一大二不分系)"), ("fGrade", "全部")],],
                     "courses.xlsx", "IPU")

# fetch_from_timetable([[("fAcySem", "114 學年度 第 2 學期"), ("fType", "學士班共同課程"), ("fCategory", "校共同課程"), ("fDep", "核心課程")]],
#                      "courses.xlsx", "1142_General")