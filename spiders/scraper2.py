import scrapy
from scrapy.utils.response import open_in_browser
import xlsxwriter
import json
from scrapy import signals



class QuotesSpider(scrapy.Spider):
    name = "shamu"
    shamuKiList = []

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
        for i in data2:
            image = i["p_img_url"]
            name = i["p_desc"]
            brand = i["p_brand"]
            category = i["tlc_n"]
            price = i["sp"]
            mrp = i["mrp"]
            variant = i["w"]

            self.shamuKiList.append([name, image, brand, category, price, mrp, variant])




    def spider_closed(self, spider):
        shamuKiWorkbook = xlsxwriter.Workbook("shamu_laya.xlsx")
        shamuKiWorksheet = shamuKiWorkbook.add_worksheet()
        shamuKiWorksheet.write('A1', "Name")
        shamuKiWorksheet.write('B1', "Image")
        shamuKiWorksheet.write('C1', "Brand")
        shamuKiWorksheet.write('D1', "Category")
        shamuKiWorksheet.write('E1', "Price")
        shamuKiWorksheet.write('F1', "MRP")
        shamuKiWorksheet.write('G1', "Variant")

        j = 2

        for i in self.shamuKiList:
            shamuKiWorksheet.write('A{}'.format(j), i[0])
            shamuKiWorksheet.write('B{}'.format(j), i[1])
            shamuKiWorksheet.write('C{}'.format(j), i[2])
            shamuKiWorksheet.write('D{}'.format(j), i[3])
            shamuKiWorksheet.write('E{}'.format(j), i[4])
            shamuKiWorksheet.write('F{}'.format(j), i[5])
            shamuKiWorksheet.write('G{}'.format(j), i[6])


            j += 1

        shamuKiWorkbook.close()




