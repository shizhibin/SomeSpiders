# 涉及动态字体混淆反爬破解
import time
import requests
import os,re
from lxml import etree
import pandas as pd
from pandas import ExcelWriter
from fontTools.ttLib import TTFont
from io import BytesIO

class Cods():
    def __init__(self):
        self.headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Connection': 'keep-alive',
            'Cookie': 'kefuCookie=2b52e19a9bc742f8a0970b9785eb0c29; __utmz=48894260.1591086682.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); _ga=GA1.3.1148237996.1591086682; _gid=GA1.3.1012094591.1591086683; key=%5B%7B%22title%22%3A%22%E5%B7%9D%E6%B8%9D%E5%90%88%E4%BD%9C%E7%A4%BA%E8%8C%83%E5%8C%BA%E5%B9%BF%E5%AE%89%E5%8D%8F%E5%85%B4%E7%94%9F%E6%80%81%E6%96%87%E5%8C%96%E6%97%85%E6%B8%B8%E5%9B%AD%E5%8C%BA%E4%BA%BA%E5%8A%9B%E8%B5%84%E6%BA%90%E5%92%8C%E7%A4%BE%E4%BC%9A%E4%BF%9D%E9%9A%9C%E5%B1%80%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E5%B7%9D%E6%B8%9D%E5%90%88%E4%BD%9C%E7%A4%BA%E8%8C%83%E5%8C%BA%E5%B9%BF%E5%AE%89%E5%8D%8F%E5%85%B4%E7%94%9F%E6%80%81%E6%96%87%E5%8C%96%E6%97%85%E6%B8%B8%E5%9B%AD%E5%8C%BA%E4%BA%BA%E5%8A%9B%E8%B5%84%E6%BA%90%E5%92%8C%E7%A4%BE%E4%BC%9A%E4%BF%9D%E9%9A%9C%E5%B1%80%22%2C%20%22other%22%3A%22%22%7D%2C%7B%22title%22%3A%22%E5%8C%97%E4%BA%AC%E5%B0%9A%E7%9D%BF%E9%80%9A%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E5%8C%97%E4%BA%AC%E5%B0%9A%E7%9D%BF%E9%80%9A%22%2C%20%22other%22%3A%22%22%7D%2C%7B%22title%22%3A%22%E5%BE%B7%E5%AE%89%E5%8E%BF%E5%9F%8E%E4%B9%A1%E5%BB%BA%E8%AE%BE%E5%B1%80%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E5%BE%B7%E5%AE%89%E5%8E%BF%E5%9F%8E%E4%B9%A1%E5%BB%BA%E8%AE%BE%E5%B1%80%22%2C%20%22other%22%3A%22%22%7D%2C%7B%22title%22%3A%22%E6%B2%B3%E5%8D%97%E8%81%8C%E4%B8%9A%E6%8A%80%E6%9C%AF%E5%AD%A6%E9%99%A2%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E6%B2%B3%E5%8D%97%E8%81%8C%E4%B8%9A%E6%8A%80%E6%9C%AF%E5%AD%A6%E9%99%A2%22%2C%20%22other%22%3A%22%22%7D%2C%7B%22title%22%3A%22%E7%94%98%E8%82%83%E5%8D%8E%E6%98%8E%E7%94%B5%E5%8A%9B%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E7%94%98%E8%82%83%E5%8D%8E%E6%98%8E%E7%94%B5%E5%8A%9B%E8%82%A1%E4%BB%BD%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8%22%2C%20%22other%22%3A%22%22%7D%2C%7B%22title%22%3A%22%E5%AE%89%E5%BA%B7%E5%B8%82%E4%BA%BA%E6%B0%91%E5%8C%BB%E9%99%A2%22%2C%20%22link%22%3A%22wx_searchPro.action%3Fkeyword%3D%E5%AE%89%E5%BA%B7%E5%B8%82%E4%BA%BA%E6%B0%91%E5%8C%BB%E9%99%A2%22%2C%20%22other%22%3A%22%22%7D%5D; Hm_lvt_f4e96f98fa73da7d450a46f37fffbf56=1591086682,1591157922; __utma=48894260.1148237996.1591086682.1591086682.1591157923.2; __utmc=48894260; __utmt=1; __utmb=48894260.2.10.1591157923; Hm_lpvt_f4e96f98fa73da7d450a46f37fffbf56=1591157947; JSESSIONID=D4119DD0F04A4E8C0EA46FA322320C59; userCookie=612b193d-60b2-4470-2b54-6072e57f2718',
            'Host': 'ss.cods.org.cn',
            'Referer': 'https://ss.cods.org.cn/latest',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
        }
        self.base_path = os.path.dirname(os.path.abspath(__file__))

    def get_number(self,company_name):
        res = requests.get("https://ss.cods.org.cn/latest/searchR?q={}&currentPage=1&t=common&searchToken=".format(company_name),headers=self.headers)
        woff_name = re.findall(r"\/(\d+\.woff2)",res.text)[0]
        # print(woff_name)
        # 获得动态字体地址并请求获得字体对象
        woff_url = "https://ss.cods.org.cn/css/woff/{}".format(woff_name)
        woff_res = requests.get(woff_url,headers=self.headers)
        font = TTFont(BytesIO(woff_res.content))
        cmap = font.getBestCmap()

        fmap = {
            0:None,1:None,2:None,3:'0',4:'1',5:'2',6:'3',7:'4',8:'5',9:'6',10:'7',11:'8',12:'9',13:'A',14:'B',15:'C',16:'D',17:'E',18:'F',19:'G',20:'H',21:'I',22:'J',23:'K',24:'L',25:'M',26:'N',27:'O',28:'P',29:'Q',30:'R',31:'S',32:'T',33:'U',34:'V',35:'W',36:'X',37:'Y',38:'Z',39:'0',40:'1',41:'2',42:'3',43:'4',44:'5',45:'6',46:'7',47:'8',48:'9',
            49:'A',50:'B',51:'C',52:'D',53:'E',54:'F',55:'G',56:'H',57:'I',58:'J',59:'K',60:'L',61:'M',62:'N',63:'O',64:'P',65:'Q',66:'R',67:'S',68:'T',69:'U',70:'V',71:'W',72:'X',73:'Y',74:'Z'
        }
        html = etree.HTML(res.text)
        number = html.xpath("//div[@class='result result-2']//div[@class='info']/h6[text()='统一社会信用代码：']/following-sibling::p[1]/text()")[0]
        # print(number)
        result = ""
        for n in number:
            n = ord(n)

            _id = font.getGlyphID(cmap[n])
            # 根据索引id获得最终结果
            r_n = fmap[_id]
            result += r_n
        return result
            
    def run(self):
        for _,_,files in os.walk(self.base_path):
            for f in files:
                if not f.startswith('~') and f.endswith('xlsx'):
                    ex = pd.read_excel(r"{}\{}".format(self.base_path,f))
                    df = pd.DataFrame(ex)
                    numbers = []
                    for row in df.itertuples():
                        company_name = row[1]
                        try:
                            number = self.get_number(company_name)
                        except IndexError:
                            number = None
                        time.sleep(5)
                        print('公司名称:{}  统一社会信用代码:{}'.format(company_name,number))
                        numbers.append(number)
                    df['统一社会信用代码'] = numbers
                    writer = ExcelWriter(r"{}\已处理-{}".format(self.base_path,f),engine='xlsxwriter',options={'strings_to_urls':False})
                    df.to_excel(writer,index=False)
                    writer.save()


if __name__ == '__main__':
    c = Cods()
    c.run()
