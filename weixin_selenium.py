from selenium import webdriver
import time,re
from lxml import etree
from lxml.html import tostring
from datetime import datetime
from jieba.analyse import extract_tags
import pymysql

# MYSQL参数
MYSQL_HOST = ''
MYSQL_PORT = 0
MYSQL_USER_NAME = ''
MYSQL_PASSWORD = ''
MYSQL_DB = ''
MYSQL_TABLE = ''

chrome_options = webdriver.ChromeOptions()
# 使用headless无界面浏览器模式
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--disable-gpu')

class WeixinSpder():
    def __init__(self,target=None,sleep=1,headless=False,page=10):
        if not target:
            self.targets = []
        self.TIME_SLEEP = sleep
        self.browser = webdriver.Chrome()
        self.headless = headless
        if self.headless:
            self.browser = webdriver.Chrome(chrome_options=chrome_options)
        self.max_page = page
        self.targets = target
        self.conn = pymysql.connect(host=MYSQL_HOST,port=MYSQL_PORT,user=MYSQL_USER_NAME,password=MYSQL_PASSWORD,db=MYSQL_DB)
        self.cur = self.conn.cursor()

    def search(self,search_keyword):
        print('开始搜索..')
        self.browser.get('https://www.sogou.com/')
        time.sleep(self.TIME_SLEEP)
        self.browser.get('https://v.sogou.com/?p=40020600&query=&ie=utf8')
        time.sleep(self.TIME_SLEEP)
        self.browser.get('https://weixin.sogou.com/')
        time.sleep(self.TIME_SLEEP)
        query = self.browser.find_element_by_id('query')
        query.send_keys(search_keyword)
        time.sleep(self.TIME_SLEEP)
        swz = self.browser.find_element_by_class_name('swz')
        swz.click()
        time.sleep(self.TIME_SLEEP)
        r = self.parse_list_page(search_keyword)
        if r == 'success':
            time.sleep(self.TIME_SLEEP)
        else:
            self.browser.quit()
            self.browser = webdriver.Chrome()


    def parse_list_page(self,search_keyword):
        page_count = 1
        while page_count <= self.max_page:
            print('解析列表页..',"当前为第{}页".format(page_count))
            titles = self.browser.find_elements_by_xpath("//a[contains(@uigs,'article_title')]")
            list_page_handle = self.browser.current_window_handle
            for title in titles:
                title.click()
                handles = self.browser.window_handles
                for newhandle in handles:
                    if newhandle != list_page_handle:
                        self.browser.switch_to.window(newhandle)
                        self.parse_detail_page(search_keyword)
                        self.browser.close()
                self.browser.switch_to.window(list_page_handle)
            time.sleep(self.TIME_SLEEP)
            # 翻页
            try:
                next_button = self.browser.find_element_by_class_name('np')
            except:
                print('异常翻页 跳过...')
                return 'error'
            next_button.click()
            page_count += 1
            time.sleep(self.TIME_SLEEP)
        return 'success'

    def parse_detail_page(self,search_keyword):
        print('解析详细页面..')
        time.sleep(self.TIME_SLEEP)
        html = etree.HTML(self.browser.page_source)
        item = {}
        item['url'] = self.browser.current_url
        item['title'] = self.browser.title
        try:
            item['source'] = html.xpath('//a[@id="js_name"]/text()')[0].strip()
        except:
            item['source'] = None
        item['target'] = search_keyword
        try:
            publish_time = re.findall(r'var t=\"\d+\",n=\"\d+\",s=\"(.*?)\"',self.browser.page_source)[0]
        except:
            print('页面异常  跳过')
            return
        item['publish_time'] = datetime.strptime(publish_time, '%Y-%m-%d')
        content_html = html.xpath("//div[@id='js_content']")[0]
        content_html = tostring(content_html,encoding='utf-8').decode('utf-8')
        item['content_html'] = content_html
        item['content'] = self.browser.find_element_by_id("js_content").text
        keyword = extract_tags(item['content'],topK=2, allowPOS=('n','v','vn'))
        item['keyword'] = ";".join(keyword)
        item['init_time'] = datetime.now()
        item['category'] = 'weixin'
        self.insert_sql(item)
    
    def insert_sql(self,item):
        keys = ','.join(item.keys())
        values = ','.join(['%s'] * len(item))
        sql = 'INSERT INTO {table} ({keys}) values ({values})'.format(table=MYSQL_TABLE, keys=keys, values=values)
        try:
            self.cur.execute(sql, tuple(item.values()))
            self.conn.commit()
            print('数据入库成功')
        except Exception as e:
            print('##################Failed#################', e)
            self.conn.rollback()

    def run(self):
        for target in self.targets:
            self.search(target)



if __name__ == '__main__':
    target = [
        "知远战略与防务研究所","工信微报","军工圈","海洋防务前沿","TechWeb","半导体行业观察","云计算头条","工业大数据","数博会","财政资金申请","中国信息化百人会",
        "极客公园","Analysys易观"
        ]
    spider = WeixinSpder(target,sleep=2,page=20)
    spider.run()



