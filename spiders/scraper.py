import scrapy
from scrapy.utils.response import open_in_browser
import xlsxwriter
from scrapy import signals



class QuotesSpider(scrapy.Spider):
    name = "ramu"
    ramuKaSet = set()
    ramuKiList = []

    @classmethod
    def from_crawler(cls, crawler, *args, **kwargs):
        spider = super(QuotesSpider, cls).from_crawler(crawler, *args, **kwargs)
        crawler.signals.connect(spider.spider_closed, signal=signals.spider_closed)
        return spider

    def start_requests(self):
        url = "https://www.bigbasket.com/cl/cleaning-household/?nc=nb#!page={}"
        for i in range(1,10):
            yield scrapy.Request(url=url.format(i), callback=self.parse)


    def parse(self, response):
        table = response.xpath('//*[@id="products-container"]//li[@qa="product"]')
        for i in table:
            image = i.xpath('.//a/img[@src]').extract()[0].split(' data-src="//')[1].split('">')[0]
            if image in self.ramuKaSet:
                continue
            else:
                self.ramuKaSet.add(image)
            name = i.xpath('.//span[@qa="prodNameRP"]//text()').extract()[0].replace("...","")
            price = i.xpath('.//div[@class="uiv2-rate-count-avial"]//text()').extract()[1].strip()
            brand = i.xpath('.//span[@class="uiv2-brand-title"]//text()').extract()[0]
            self.ramuKiList.append([image, name, price, brand])


    def spider_closed(self, spider):
        ramuKiWorkbook = xlsxwriter.Workbook("ramu_laya.xlsx")
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




