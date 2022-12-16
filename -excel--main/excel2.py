import json
import urllib.parse

import requests
from bs4 import BeautifulSoup
import pandas as pd


class CertInfoScraper:
    def __init__(self, name, certnum):
        self.name = name
        self.certnum = certnum
        self.login_url = "http://cx.mem.gov.cn/"
        self.data_url = "http://cx.mem.gov.cn/cms/html/certQuery/certQuery.do"
        Hm_lvt = '1669365331,1669422090,1671070277'
        acw_tc = '7b39759016711251314186195e44ad40c43614c37ae50b7951825feb6d80d8'
        JSESSIONID = 'F8E7D0F907CFF9F119B1E13E47DB8848'
        self.headers = {
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
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 "
                          "Safari/537.36",
            "X-Requested-With": "XMLHttpRequest"
        }
        self.cookies = {
            'Hm_lvt_7c3492d683dc7a90fd44bf8bfd57e50c': Hm_lvt,
            'acw_tc': acw_tc,
            'JSESSIONID': JSESSIONID,
        }

    def login(self):
        login_response = requests.post(self.login_url, cookies=self.cookies)

    def get_session_id(self):
        params = {
            "method": "getServerTime"
        }
        data_response = requests.post(self.data_url, headers=self.headers, params=params)
        response_content = data_response.text
        response_dict = json.loads(response_content)
        return response_dict["time"]

    def scrape(self):
        encoded_name = urllib.parse.quote(self.name)
        encoded_name = urllib.parse.quote(encoded_name)
        session_id = self.get_session_id()
        data_html = requests.get(self.data_url + "?"
                                                 "method=getCertQueryResult"
                                                 "&ref=ch&certtype_code=720"
                                                 "&certnum=" + self.certnum +
                                                "&stu_name=" + encoded_name +
                                                "&passcode=1345"
                                                "&sessionId=" + session_id)
        html_soup = BeautifulSoup(data_html.text, "html.parser")
        tables = html_soup.find_all('table')
        for i, table in enumerate(tables):
            df = pd.read_html(str(table), flavor='html5lib')[0]
            df.to_excel('data.xlsx', index=False, startrow=i, startcol=0)

def main():
    scraper = CertInfoScraper("韩泷沣", "23900519961004251x")
    scraper.login()
    scraper.scrape()

main()
#Add GUI for this program