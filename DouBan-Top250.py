#爬取豆瓣电影Top250列表数据程序。模拟浏览器访问网站，抓取网页上的数据。
#电影详情链接、图片链接、中文名、外国名、评分、评价人数、简介等，并保存到Excel文件中。

# -*- coding = utf-8 -*-
from bs4 import BeautifulSoup
import re
import requests
import xlwt
import time
time.sleep(2)  #每次请求后等待2秒.模拟人类用户的访问。

findLink = re.compile(r'<a href="(.*?)">')
findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S)
findTitle = re.compile(r'<span class="title">(.*)</span>')
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
findJudge = re.compile(r'<span>(\d*)人评价</span>')
findInq = re.compile(r'<span class="inq">(.*)</span>')
findBd = re.compile(r'<p class="">(.*?)</p>', re.S)

def main():
    baseurl = "https://movie.douban.com/top250?start=" 
    datalist = getData(baseurl)
    savepath = "豆瓣电影Top250.xls"
    saveData(datalist, savepath)  

#爬取网页
def getData(baseurl):
    datalist = []  
    for i in range(0, 10):
        url = baseurl + str(i * 25)
        html = askURL(url)
        if html:  
            soup = BeautifulSoup(html, "html.parser")
            for item in soup.find_all('div', class_="item"):
                data = []
                item = str(item)
                link = re.findall(findLink, item)[0]
                data.append(link)
                imgSrc = re.findall(findImgSrc, item)[0]
                data.append(imgSrc)
                titles = re.findall(findTitle, item)
                if len(titles) == 2:
                    ctitle = titles[0]
                    data.append(ctitle)
                    otitle = titles[1].replace("/", "")
                    data.append(otitle)
                else:
                    data.append(titles[0])
                    data.append(' ')
                rating = re.findall(findRating, item)[0]
                data.append(rating)
                judgeNum = re.findall(findJudge, item)[0]
                data.append(judgeNum)
                inq = re.findall(findInq, item)
                if len(inq) != 0:
                    inq = inq[0].replace("。", "")
                    data.append(inq)
                else:
                    data.append(" ")
                bd = re.findall(findBd, item)[0]
                bd = re.sub('<br(\s+)?/>(\s+)?', "", bd)
                bd = re.sub('/', "", bd)
                data.append(bd.strip())
                datalist.append(data)
    return datalist

def askURL(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=head)
        response.raise_for_status()  
        return response.text  
    except requests.exceptions.RequestException as e:
        print(f"Request Error: {e}")
        return None

#保存数据到表格
def saveData(datalist, savepath):
    print("save.......")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)
    col = ("电影详情链接", "图片链接", "影片中文名", "影片外国名", "评分", "评价数", "概况", "相关信息")
    for i in range(0, 8):
        sheet.write(0, i, col[i])
    for i in range(0, 250):
        data = datalist[i]
        for j in range(0, 8):
            sheet.write(i + 1, j, data[j])
    book.save(savepath)

if __name__ == "__main__":
    main()
    print("爬取完毕！")




