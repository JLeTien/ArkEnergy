import scrapy


class ExampleSpider(scrapy.Spider):
    name = "example"
    allowed_domains = ["https://www.energymagazine.com.au/"]
    start_urls = ["https://www.energymagazine.com.au/renewable-energy/"]

    def parse(self, response):
        news = response.css('div.post')
        post_contents =  news.css('div.post-content')
        post_headers = [post_content.xpath('./div[@class="post-header"]') for post_content in post_contents]
        #print(post_headers)
        
        
        print("===========================================================================")
        for post_content in post_contents:
            # Use XPath to select the <a> tag within the <div class="post-header">
            a_tag = post_content.xpath('./div[@class="post-header"]/h3/a')

            # Extract the text content of the <a> tag
            if a_tag:
                text = a_tag.xpath('text()').extract_first()
                print(text)
            else:
                print("No <a> tag found within the post-header.")
        print("===========================================================================")