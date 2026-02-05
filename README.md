# CourseHelper

這是一系列的自動化工具，幫助自動規劃課程、自動下載課表......等功能

## 功能
* ~~自動規劃課程（開發中）~~
* 自動下載課表

## 特色
* 自動化操作：使用 Selenium 模擬瀏覽器行為，自動填選學年、學院、系所等條件。
* API 攔截：不依賴傳統的 HTML 解析，而是透過 JavaScript 攔截後端傳回的 JSON 資料，確保資料準確性與完整性。
* Excel 輸出：自動比對並過濾重複課程，將結果分頁籤（Sheet）儲存為 courses.xlsx。
* 進度顯示：使用 tqdm 顯示抓取與寫入進度。

## 安裝指南

### 環境要求

* Python 3.8 或以上版本
* Google Chrome 瀏覽器 (程式會自動下載對應的驅動程式)

### 安裝步驟

1. Clone repository

```bash
git clone https://www.github.com/alpha34727/CourseHelper
```

2. Install requirements

```bash
pip install -r requirements.txt
```

3. Run

```bash
python ./get_course.py
```

## 使用說明

```python
fetch_from_timetable(req: List[List[(str, str)]], filename: str, sheetname: str)
```

### reqs (Requirements)
× 這是最複雜的參數，它是一個列表 (List)，裡面包含了多次查詢的條件。
* 結構：[[查詢1條件], [查詢2條件], ...]
* 用途：例如你想在同一個 Excel 分頁（如 "BME"）中，同時存入「上學期」和「下學期」的課，就可以在 reqs 裡放兩組條件。

#### 查詢條件格式

```python
(下拉選單id, 要選的文字）
```

### filename
輸出的 Excel 檔名（例如 "courses.xlsx"）。

### sheetname
這些資料要存在 Excel 的哪一個分頁（例如 "MT" 或 "BME"）。
