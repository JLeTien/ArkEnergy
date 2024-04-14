import scrapy
from scrapy.crawler import CrawlerProcess
from scrapy_splash import SplashRequest

# scraping based on keys words = "wind/rebate/...."
# clean energy council / government run energy council / renewable economy
    
class Spider1(scrapy.Spider):
    name = "example"
    allowed_domains = ["energymagazine.com.au"]
    start_urls = ["https://www.energymagazine.com.au/renewable-energy/"]
    slides_data = []

    def parse(self, response):
        # Extracting information from post headers
        news = response.css('div.post')
        post_contents =  news.css('div.post-content')
        
        # for post_content in post_contents:
        #     # Use XPath to select the <a> tag within the <div class="post-header">
        #     a_tag = post_content.xpath('./div[@class="post-header"]/h3/a')
        #     if a_tag:
        #         text = a_tag.xpath('text()').extract_first()
        #         print(text)
        #     else:
        #         print("No <a> tag found within the post-header.")

        post_headers = response.css('div.post-header')
        for header in post_headers[:3]:
            # Extracting the URL from the <a> tag within the post header
            title = header.css('h3.post-title.entry-title a::text').extract_first()
            # print("title:" + title)
            
            url = header.css('h3.post-title.entry-title a::attr(href)').extract_first()
            # print("link: " + url)

            # Need to use SpashRequest to handle dynamically rendered pages in JS
            yield SplashRequest(url, self.parse_post, args={'wait': 2}, meta={'title': title})

    def parse_post(self, response):
        paragraphs = response.xpath('//div[@class="article-content"]/p')
        content = ''

        for p in paragraphs:
            paragraph_text = p.xpath('string()').get().strip() 
            content += paragraph_text  
        
        title = response.meta['title']
        self.slides_data.append({'title': title, 'content': content})
        # print(content)
    
class Spider2(scrapy.Spider):
    name = "example2"
    allowed_domains = ["reneweconomy.com.au"]
    start_urls = ["https://reneweconomy.com.au/all-articles/"]
    slides_data = []
    
    def parse(self, response):
        print("====================================================================")
        h2_element = response.css('h2.wp-block-post-title')
        for header in h2_element[:3]:

            title = header.css('h2.wp-block-post-title a::text').get()
            
            url = header.css('a::attr(href)').extract_first()
            # print("link: " + url)
            
            yield SplashRequest(url, self.parse_post, args={'wait': 2}, meta={'title': title})
        print("====================================================================")
       
    def parse_post(self, response):
        paragraphs = response.xpath('//div[contains(@class, "entry-content") '
                                    'and contains(@class, "wp-block-post-content") '
                                    'and contains(@class, "is-layout-flow") '
                                    'and contains(@class, "wp-block-post-content-is-layout-flow")]//p')
        
        content = ''

        for p in paragraphs:
            paragraph_text = p.xpath('string()').get().strip() 
            content += paragraph_text
        # print(content)
        
        title = response.meta['title']
        self.slides_data.append({'title': title, 'content': content})
        # print(self.slides_data)
        