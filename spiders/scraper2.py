import scrapy
from scrapy.utils.response import open_in_browser
import xlsxwriter
import json
from scrapy import signals



class QuotesSpider(scrapy.Spider):
    name = "shamu"
    ramuKaSet = set()
    ramuKiList = []

    @classmethod
    def from_crawler(cls, crawler, *args, **kwargs):
        spider = super(QuotesSpider, cls).from_crawler(crawler, *args, **kwargs)
        crawler.signals.connect(spider.spider_closed, signal=signals.spider_closed)
        return spider

    def start_requests(self):
        url = "https://www.bigbasket.com/product/get-products/?slug=beverages&page={}&tab_type=[%22all%22]&listtype=pc"
        for i in range(2,41):
            yield scrapy.Request(url=url.format(i), callback=self.parse)


    def parse(self, response):
        data = json.loads(response.body)
        data2 = data["tab_info"]["product_map"]["all"]["prods"]

        print(data2)




    def spider_closed(self, spider):
        ramuKiWorkbook = xlsxwriter.Workbook("shamu_laya.xlsx")
        ramuKiWorksheet = ramuKiWorkbook.add_worksheet()
        ramuKiWorksheet.write('A1', "Image")
        ramuKiWorksheet.write('B1', "Name")
        ramuKiWorksheet.write('C1', "Price")
        ramuKiWorksheet.write('D1', "Brand")

        j = 2

        for i in self.ramuKiList:
            ramuKiWorksheet.write('A{}'.format(j), i[0])
            ramuKiWorksheet.write('B{}'.format(j), i[1])
            ramuKiWorksheet.write('C{}'.format(j), i[2])
            ramuKiWorksheet.write('D{}'.format(j), i[3])
            j += 1

        ramuKiWorkbook.close()




