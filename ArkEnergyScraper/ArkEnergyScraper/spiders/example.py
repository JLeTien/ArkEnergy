import scrapy
from scrapy.crawler import CrawlerProcess
from scrapy_splash import SplashRequest

# scraping based on keys words = "wind/rebate/...."
# clean energy council / government run energy council / renewable economy
    
class ExampleSpider(scrapy.Spider):
    name = "example"
    allowed_domains = ["energymagazine.com.au"]  # Removed the protocol and path
    start_urls = ["https://www.energymagazine.com.au/renewable-energy/"]
    
    def start_requests(self):
        for url in self.start_urls:
            yield SplashRequest(url, self.parse,
                args={'wait': 2})  # Adjust the wait time as needed

    def parse(self, response):
        # Extracting information from post headers
        post_headers = response.css('div.post-header')
        news = response.css('div.post')
        post_contents =  news.css('div.post-content')
        
        print("===========================================================================")
        
        for post_content in post_contents:
            # Use XPath to select the <a> tag within the <div class="post-header">
            a_tag = post_content.xpath('./div[@class="post-header"]/h3/a')
            if a_tag:
                text = a_tag.xpath('text()').extract_first()
                print(text)
            else:
                print("No <a> tag found within the post-header.")
        
        print("===========================================================================")
        
        post_headers = response.css('div.post-header')
        for header in post_headers[:1]:
            # Extracting the URL from the <a> tag within the post header
            url = header.css('h3.post-title.entry-title a::attr(href)').extract_first()
            print("link: " + url)
            # yield scrapy.Request(url, callback=self.parse_post)
            
            yield SplashRequest(url, self.parse_post,
                args={'wait': 2})  # Adjust the wait time as needed
            
        print("===========================================================================")

    def parse_post(self, response):
        content = response.xpath('//div[@class="article-content"]/*[not(@class="addtoany_share_save_container addtoany_content addtoany_content_top")]').get()

        print(content)
        
        # yield {
        #     'title': title,
        #     'content': content
        # }

def main():
    process = CrawlerProcess()
    process.crawl(ExampleSpider)
    process.start()  # the script will block here until the crawling is finished

if __name__ == "__main__":
    main()