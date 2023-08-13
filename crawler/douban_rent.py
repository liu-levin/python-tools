import time
import os
import re
import random
from datetime import datetime, timedelta
from dateutil import tz
import requests
import xlsxwriter as xw


RentInfoFileName = "./rentInfo.xlsx"
MaxPage = 5

def local_time(local_tz="Asia/Shanghai"):
    utc = datetime.utcnow()  # naive datetime object
    utc = utc.replace(tzinfo=tz.gettz("UTC"))  # timezone aware datetime object

    # convert time zone
    local = utc.astimezone(tz.gettz(local_tz))
    return local



def crawlGroup(group_id, group_name, cookie=None):
    print("=========================================================")
    print(f"开始爬取豆瓣小组：【{group_name}】\n")
    print(f"url = https://www.douban.com/group/{group_id}/\n")
    print("开始执行，时间：", datetime.strftime(local_time(), "%Y-%m-%d %H:%M:%S") + "\n")

    page_infos = []
    for page in range(1, MaxPage + 1):  # 从最新页开始爬取
        print(" * ", end="")  # 显示爬取进度
        page_info = crawlPage(group_id, page,cookie)  # 爬取一页
        page_infos.extend(page_info)
        randomSleep()  # 爬取不要太快，防止被封
        infoLength = len(page_info)
        if infoLength > 0:
           item = page_info[infoLength-1]
           if len(item) == 5:
               if not isValidTime(item[3]):
                   break           
    return page_infos


def crawlPage(doubanGroupId, page, cookie=None):

    url = f"https://www.douban.com/group/{doubanGroupId}/discussion?start={25*(page-1)}&type=new"
    if cookie is None:
        cookie = os.environ.get("COOKIE")

    headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
    "Cookie": cookie,
    }

    # raw text
    html_raw = requests.get(url, headers=headers).text

    # raw items
    items_raw = re.findall(r"<tr class=\"\">([\s\S]*?)<\/tr>", html_raw)
    print("items length: ",len(items_raw))

    page_info = []
    for i in range(len(items_raw)):
        item = items_raw[i]
        item = re.sub(" +", " ", item.replace("\n", " "))
        link_title = re.findall(r"<a href=\"(.*?)\"\s+title=\"(.*?)\"", item)
        author = re.findall(r"<a href=\".*?\" class=\"\">(.*?)<\/a>", item)
        respCount = re.findall(r"class=\"r-count \">(.*?)<\/td>", item)
        update = re.findall(r"class=\"time\">(.*?)<\/td>", item)
        title = link_title[0][1]
        link = link_title[0][0]
        timestamp = update[0]

        poster = ""
        if len(author) > 0:
            poster = author[1]

        count = 0
        if len(respCount) > 0:
            count = respCount[0]
        # print(title,poster,count,timestamp,link,"\n")
        page_info.append([title,poster,count,timestamp,link])

    return page_info

    

def filterInfo(title):
    if "求租" in title:
        return True
    
    if "转租" in title:
        return False
    
    keywords = [
        "整租",
    ]

    for keyword in keywords: 
        if keyword in title:
            return False
        
    return False 

def filterDuplication(records):
    newRecords = []
    s = set()
    for record in records:
        if len(record) == 5:
            if not (record[4] in s):
                s.add(record[4])
                newRecords.append(record)
    return newRecords



def crawlGroupFollowerCount(groups):
    for groupId,groupName in groups:
        url = f"https://www.douban.com/group/{groupId}/"
        cookie = getCookie()
        headers = getHeader(cookie)
        # raw text
        html_raw = requests.get(url, headers=headers).text
        # randomSleep()
        # raw items
        content = re.findall(r"<a href=\".*?\">浏览所有成员(.*?)</a>", html_raw)
        if len(content) > 0:
            count = content[0]
            count = count.strip(' ').strip('(').strip(')')
            print( groupName," : ",count)





def write_to_xlsx(doubanGroups, cookie):
    colNames = ["标题","作者","回应","最后回应","链接"]
    workBook = xw.Workbook(RentInfoFileName)

    for groupId,groupName in doubanGroups:
        records = crawlGroup(groupId,groupName,cookie)
        records = filterDuplication(records)

        worksheet = workBook.add_worksheet(groupName)
        worksheet.activate()
        worksheet.set_column("A:A", 100)
        worksheet.set_column("B:B", 18)
        worksheet.set_column("C:C", 6)
        worksheet.set_column("D:D", 20)
        worksheet.set_column("E:E", 50)
        worksheet.write_row("A1", colNames)

        print(records)
        for i, row in enumerate(records):
            if len(row) !=5:
                continue
            title = row[0]
            if filterInfo(title):
                continue
            idx = "A" + str(i + 2)
            if len(row) == 5:
                worksheet.write_row(idx, row)
    
    workBook.close()    


def getDoubanGroups():
    groups = [
        # ("HZhome","杭州租房1"),
        # ("467221", "杭州租房2"),
        # ("hzhouse", "杭州租房3"),
        # ("623520","杭州租房4"),
        # ("501627","杭州租房5"),
        # ("606541","杭州租房6"),
        # ("578360","杭州租房7"),
        # ("224803","我要在杭州租房子"),
        # ("276209","杭州租房一族"),
        # ("560075","杭州西湖区租房"),
        # ("551531","杭州租房大全"),
        # ("145219","杭州 出租 租房 中介免入"),
        # ("120199","共享天堂---我要租房"),
        # ("257587","杭州租房小组"),
        # ("576850","杭州豆瓣租房"),
        ("595637","杭州滨江萧山租房"),
        ("550725","滨江萧山租房"),
        # ("550728","杭州滨江 租房 整 合 拼 转"),
        # ("598372","杭州萧山租房"),
        # ("554566","杭州滨江租房"),
        # ("HZ_home","杭州滨江租房"),
        ("568935","杭州城西租房")
    ]
    return groups




def randomSleep():
    num = random.randint(10,19)
    time.sleep(num)

def getCookie():
    return ""
def getHeader(cookie):
    headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Cookie": cookie,
        }
    return headers

def isValidTime(ts):

    checkTime =  datetime.now() + timedelta(days=10)
    expiredTime = datetime.now() - timedelta(days=10)

    yearNum = datetime.now().year

    datetime_str = str(yearNum)+"-"+ts
    articleTime = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M')

    if articleTime > checkTime:
        datetime_str = str(yearNum-1)+"-"+ts
        articleTime = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M')
    return articleTime>expiredTime
    

def main():
    doubanGroups = getDoubanGroups()
    write_to_xlsx(doubanGroups,getCookie())
    # crawlGroupFollowerCount(doubanGroups)
       
if __name__ == '__main__':
    main()

