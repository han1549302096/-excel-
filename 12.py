import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import urllib.parse

name = "刘岩涛"
code1 = "371083199204154017"
encoded_str = urllib.parse.quote(name)
encoded_str = urllib.parse.quote(encoded_str)
print(encoded_str)
url = "http://cx.mem.gov.cn/"
url1 = "http://cx.mem.gov.cn//cms/html/certQuery/certQuery.do?method=getServerTime"
Hm_lvt = '1669365331,1669422090,1671070277'
acw_tc = '7b39759016711251314186195e44ad40c43614c37ae50b7951825feb6d80d8'
JSESSIONID = 'F8E7D0F907CFF9F119B1E13E47DB8848'
# 设置 cookie
cookies = {
    'Hm_lvt_7c3492d683dc7a90fd44bf8bfd57e50c': Hm_lvt,
    'acw_tc': acw_tc,
    'JSESSIONID': JSESSIONID,
}

response = requests.post(url, cookies=cookies)
# 设置请求头
headers = {
    "Accept": "text/plain, */*; q=0.01",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9,zh-TW;q=0.8,en-US;q=0.7,en;q=0.6",
    "Connection": "keep-alive",
    "Content-Length": "0",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cookie": Hm_lvt + acw_tc + JSESSIONID,
    "Host": "cx.mem.gov.cn",
    "Origin": "http://cx.mem.gov.cn",
    "Referer": "http://cx.mem.gov.cn/cms/html/certQuery/certQuery.do?method=getCertQueryIndex&ref=ch",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest"
}

# 设置请求参数
params = {
    "method": "getServerTime"
}

# 发送POST请求
url = "http://cx.mem.gov.cn/cms/html/certQuery/certQuery.do"
response1 = requests.post(url, headers=headers, params=params)

# 获取响应内容
content = response1.text
response_dict = json.loads(content)
print(response_dict["time"])  # 输出响应字典
print(content)  # 输出响应内容
html = requests.get("http://cx.mem.gov.cn/cms/html/certQuery/certQuery.do?"
                   "method=getCertQueryResult&"
                   "ref=ch&certtype_code=720&"
                   "certnum="+code1+"&"
                   "stu_name="+encoded_str+"&"
                   "passcode=1345&"
                   "sessionId="+response_dict["time"])
soup = BeautifulSoup(html.text, "html.parser")
print(soup)

# 使用 find_all() 函数找到所有的 table 元素
tables = soup.find_all('table')

# 遍历每个 table 元素
for i, table in enumerate(tables):
  # 使用 read_html() 函数将表格解析为 DataFrame 对象
  df = pd.read_html(str(table), flavor='html5lib')[0]
  # 将数据写入 Excel 文件的第 i + 1 个单元格中
  df.to_excel('data.xlsx', index=False, startrow=i, startcol=0)