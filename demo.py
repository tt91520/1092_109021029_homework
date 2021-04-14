import requests
from bs4 import BeautifulSoup as bs
import openpyxl
import time
from openpyxl import Workbook
import csv
import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()

url = "https://csie.asia.edu.tw/project/{}"

def num(url, start, end):
    urllist = []
    x = str
    for i in range(start, end+1):
        if i==106:
            x="semester-1061"
            urllist.append(url.format(x))

        elif i==108:
            x="108學年"
            urllist.append(url.format(x))
        
        else:
            x="semester-{}".format(i)
            urllist.append(url.format(x))
    
    return urllist

def get_resource(url):
    headers = {"user-agent":"Mozilla/5.0(Windows NT 10.0; Win64; x64) ApplWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.329.130 Safari/537.36"}
    return requests.get(url, headers=headers,verify = False)

def webcode(r):
    return bs(r, "html.parser")

def get_word(soup,file):
    word = []
    d=[]
    re1= soup.findAll("td")
    d.append(file)  
    #print(len(re1))
    for val in re1:        
        text1=val.text.replace("\n","")
        #text2=text1.replace(" "[2],"\n")
        d.append(text1)
        #text3=text2.replace("1","\n")
        #d.append(text2)
    word.append(d)
    
    return word

def catchingwordbot(urls):
    dt=[]
    for i in urls:
        file = i.split("/")[-1]
        print("catching: ", file, "web data...")
        r = get_resource(i)
        r.encoding="utf8"
        if r.status_code == requests.codes.ok:
            soup = webcode(r.text)
            words = get_word(soup, file)  
            #print(words)
            dt = dt + words
            print("wait 5 sec")
            time.sleep(5)
        else:
            print("http request error!")

    return dt

def excelsave(word,name):
    '''wb = Workbook()
    ws = wb.active
    for i in word:
        ws.append(i)
    wb.save(f'{name}.csv')''' 
    with open(f'{name}.csv', 'w', newline='', encoding="UTF8") as csvfile:
           
        writer = csv.writer(csvfile)
        for i in word:
            # 寫入一列資料
            writer.writerow(i)



if __name__ == "__main__":
    x="projectsList"
    urls = num(url, 100, 108)
    dt = catchingwordbot(urls)
    excelsave(dt,x)
    #print(dt)
    

    