"""
按公司名称在天眼查采集疑似实际控制人信息 及总股权比例
"""
import requests
import time
import json
import os
import pandas as pd
from pandas import ExcelWriter

class TYC():
    def __init__(self):
        self.headers = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Host': 'sp0.tianyancha.com',
            'Origin': 'https://www.tianyancha.com',
            'Referer': 'https://www.tianyancha.com/?jsid=SEM-BAIDU-PZ2005-SY-000001',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-site',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
            'version': 'TYC-Web',
            'Cookie': 'TYCID=59e1e540891011ea8f2d8bbb188de0d6; undefined=59e1e540891011ea8f2d8bbb188de0d6; ssuid=2594907520; _ga=GA1.2.1736314229.1588051397; jsid=SEM-BAIDU-PZ2005-SY-000001; _gid=GA1.2.1585841997.1590648800; tyc-user-phone=%255B%252215201264869%2522%255D; RTYCID=b25c805b0e2c41f5b069c8c70102cb4b; CT_TYCID=a62ac013d49442a4a344c6532adafb9b; aliyungf_tc=AQAAAI9TsCLFhgEAgsn52jS4GkdfKeBO; csrfToken=mQuNFxqwjb4EN2QLYxhN9Qvt; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1590648800,1590713479; bannerFlag=true; tyc-user-info=%257B%2522claimEditPoint%2522%253A%25220%2522%252C%2522vipToMonth%2522%253A%2522false%2522%252C%2522explainPoint%2522%253A%25220%2522%252C%2522integrity%2522%253A%252210%2525%2522%252C%2522state%2522%253A%25225%2522%252C%2522surday%2522%253A%252233%2522%252C%2522announcementPoint%2522%253A%25220%2522%252C%2522bidSubscribe%2522%253A%2522-1%2522%252C%2522vipManager%2522%253A%25220%2522%252C%2522onum%2522%253A%252264%2522%252C%2522monitorUnreadCount%2522%253A%2522126%2522%252C%2522discussCommendCount%2522%253A%25221%2522%252C%2522claimPoint%2522%253A%25220%2522%252C%2522token%2522%253A%2522eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxNTIwMTI2NDg2OSIsImlhdCI6MTU5MDcxMzc3OSwiZXhwIjoxNjIyMjQ5Nzc5fQ.3oWJcRmcNCYmmgtogygANJpWl5y7ywk87vHRdpu_riToOufQiNP8ITk5e1YaiawaePOzHjnCsWFui8Kr0yQXIw%2522%252C%2522vipToTime%2522%253A%25221593498183942%2522%252C%2522redPoint%2522%253A%25220%2522%252C%2522myAnswerCount%2522%253A%25220%2522%252C%2522myQuestionCount%2522%253A%25220%2522%252C%2522signUp%2522%253A%25220%2522%252C%2522nickname%2522%253A%2522%25E4%25B8%25AD%25E5%25B0%258F%25E4%25BC%2581%25E4%25B8%259A%25E5%25AE%25A4%2522%252C%2522privateMessagePointWeb%2522%253A%25220%2522%252C%2522privateMessagePoint%2522%253A%25220%2522%252C%2522isClaim%2522%253A%25220%2522%252C%2522isExpired%2522%253A%25220%2522%252C%2522pleaseAnswerCount%2522%253A%25220%2522%252C%2522bizCardUnread%2522%253A%25220%2522%252C%2522vnum%2522%253A%252220%2522%252C%2522mobile%2522%253A%252215201264869%2522%257D; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxNTIwMTI2NDg2OSIsImlhdCI6MTU5MDcxMzc3OSwiZXhwIjoxNjIyMjQ5Nzc5fQ.3oWJcRmcNCYmmgtogygANJpWl5y7ywk87vHRdpu_riToOufQiNP8ITk5e1YaiawaePOzHjnCsWFui8Kr0yQXIw; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1590715123; cloud_token=b64be046994c451887edc63765c4901b; cloud_utm=8e70a1a568ea483795ec7809b33ea253; token=d99255f63e1b4d2f8cdb335b0d452dba; _utm=a5a56c707b004afebe5abe9c8fcc9634',
            'X-Auth-Token': 'eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxNTIwMTI2NDg2OSIsImlhdCI6MTU5MDcxMzc3OSwiZXhwIjoxNjIyMjQ5Nzc5fQ.3oWJcRmcNCYmmgtogygANJpWl5y7ywk87vHRdpu_riToOufQiNP8ITk5e1YaiawaePOzHjnCsWFui8Kr0yQXIw'
        }
        self.base_path = os.path.dirname(os.path.abspath(__file__))

    def get_company_id(self,company_name):
        url = "https://sp0.tianyancha.com/search/suggestV2.json"
        params = {
            'key': company_name,
            '_': int(time.time() * 1000)
        }
        res = requests.get(url, params=params, headers=self.headers)
        res = json.loads(res.text)
        try:
            return res['data'][0]['id']
        except:
            print(company_name,res)

    def search_actualcontroller(self,company_id):
        headers2 = self.headers.copy()
        headers2['Accept'] = 'application/json, text/plain, */*'
        headers2['Host'] = 'capi.tianyancha.com'
        headers2['X-AUTH-TOKEN'] = 'eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxNTIwMTI2NDg2OSIsImlhdCI6MTU5MDcxMzc3OSwiZXhwIjoxNjIyMjQ5Nzc5fQ.3oWJcRmcNCYmmgtogygANJpWl5y7ywk87vHRdpu_riToOufQiNP8ITk5e1YaiawaePOzHjnCsWFui8Kr0yQXIw'
        url = "https://capi.tianyancha.com/cloud-equity-provider/v4/actualControl/company.json?id={}&height=902&width=500".format(company_id)
        res = requests.get(url,headers=headers2)
        res = json.loads(res.text)
        actualcontroller = res['data']['actualController'].get('name')
        percent = res['data'].get('ratio')
        try:
            return actualcontroller,round(percent * 100, 2)
        except TypeError:
            return actualcontroller,None

    def run(self):
        for _,_,files in os.walk(self.base_path):
            for f in files:
                if not f.startswith('~') and f.endswith('xlsx'):
                    ex = pd.read_excel(r"{}\{}".format(self.base_path,f))
                    df = pd.DataFrame(ex)
                    actual_controllers = []
                    percents = []
                    for row in df.itertuples():
                        company_name = row[1]
                        company_id = self.get_company_id(company_name)
                        try:
                            controller,percent = self.search_actualcontroller(company_id)
                        except Exception as e:
                            controller,percent = None,None
                        print(company_name,"疑似实际控制人:",controller,"控股百分比:",percent)
                        actual_controllers.append(controller)
                        percents.append(percent)
                    df['疑似实际控制人'] = actual_controllers
                    df['控股百分比'] = percents
                    writer = ExcelWriter(r"{}\已处理-{}".format(self.base_path,f),engine='xlsxwriter',options={'strings_to_urls':False})
                    df.to_excel(writer,index=False)
                    writer.save()

if __name__ == '__main__':
    t = TYC()
    t.run()