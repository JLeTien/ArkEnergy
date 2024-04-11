from scrapy import signals
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings
from Scraper.ArkEnergyScraper.spiders.ArkSpiders import Spider1
from Scraper.ArkEnergyScraper.spiders.ArkSpiders import Spider2

def main():
    process = CrawlerProcess()
    process.crawl(Spider1)
    process.start()

if __name__ == "__main__":
    main()