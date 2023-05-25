import requests
from openpyxl import Workbook
import datetime

def get_google_rank(keywords, target_urls, api_key, cx, num_pages):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": api_key,
        "cx": cx
    }
    
    all_results = []  # 存儲所有搜索结果
    
    for keyword in keywords:
        for target_url in target_urls:
            params["q"] = keyword
            current_results = []  # 存儲當前關鍵字和目標網址的搜索结果
            
            for page in range(num_pages):
                params["start"] = page * 10 + 1  # 設置每頁的起始位置
                
                response = requests.get(url, params=params)
                data = response.json()
        
                if "items" in data:
                    for i, item in enumerate(data["items"], start=1):
                        if target_url in item["link"]:
                            rank = i + (page * 10)
                            current_results.append(rank)
            
            if current_results:
                min_rank = min(current_results)
                page = (min_rank - 1) // 10 + 1
                print(f"關鍵字 '{keyword}' 目標網址 '{target_url}' 在 Google 搜索结果中的最低排名為第 {min_rank} 名，所在頁數為第 {page} 頁")
                all_results.append((keyword, target_url, min_rank, page))
            else:
                print(f"關鍵字 '{keyword}' 目標網址 '{target_url}' 尚未在 Google 搜索结果中找到")
    
    return all_results

def create_excel_file(results):
    wb = Workbook()
    ws = wb.active
    ws.append(["Keyword 關鍵字", "Target URL 目標網址", "Rank 排名", "Page 頁數", "Timestamp 時間戳"])
    
    timestamp = datetime.datetime.now()
    for result in results:
        keyword, target_url, rank, page = result
        ws.append([keyword, target_url, rank, page, timestamp])
    
    # 调整列宽以适应较长的文本
    ws.column_dimensions["A"].width = 20  # 关键字列
    ws.column_dimensions["B"].width = 30  # 目标网址列
    ws.column_dimensions["C"].width = 20  # 排名列
    ws.column_dimensions["D"].width = 20  # 页数列
    ws.column_dimensions["E"].width = 20  # 时间戳列

    filename = f"google_rankings_{timestamp.strftime('%Y%m%d%H%M')}.xlsx"
    wb.save(filename)
    print(f"结果已保存到 {filename}")

# 输入要搜索的關鍵字列表、目標網址列表、API相關信息以及搜索的頁數
keywords = ["小琉球民宿", "小琉球星空villa", "小琉球民宿villa"]
target_urls = ["https://itravelblog.net"]
api_key = "AIzaSyC0DnMS-HPZAJsS97K2hMi2Ox5Zg_yzNrs"
cx = "e10861c636c934547"
num_pages = 10  # 設置搜索的頁數

# 獲取關鍵字在Google上的排名和頁數
results = get_google_rank(keywords, target_urls, api_key, cx, num_pages)

# 創建Excel文件並記錄排名信息
create_excel_file(results)
