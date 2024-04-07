import scrapy


class ExampleSpider(scrapy.Spider):
    name = "example"
    allowed_domains = ["https://www.energymagazine.com.au/"]
    start_urls = ["https://www.energymagazine.com.au/renewable-energy/"]

    def parse(self, response):
        # Extracting information from post headers
        post_headers = response.css('div.post-header')
        news = response.css('div.post')
        post_contents =  news.css('div.post-content')
        
        # scraping based on keys words = "wind/rebate/...."
        # clean energy council 
        # government run energy council
        # renewable economy
        
        print("===========================================================================")
        
        for post_content in post_contents:
            # Use XPath to select the <a> tag within the <div class="post-header">
            a_tag = post_content.xpath('./div[@class="post-header"]/h3/a')

            # Extract the text content of the <a> tag
            # Potentially use this to pick up on certain key words or only extract the most recent news
            if a_tag:
                text = a_tag.xpath('text()').extract_first()
                print(text)
            else:
                print("No <a> tag found within the post-header.")
        
        print("===========================================================================")
        
        # Extracting information from post headers
        post_headers = response.css('div.post-header')
        
        for header in post_headers[:3]:
            # Extracting the URL from the <a> tag within the post header
            url = header.css('h3.post-title.entry-title a::attr(href)').extract_first()
            yield {
                'url': url
            }
        
        print("===========================================================================")
        
        # def parse_post(self, response):
        # # Parse the response from the post URL
        # # Extract whatever data you need from the post page
        # # For example:
        #     title = response.css('h1.post-title::text').get()
        #     content = response.css('div.post-content::text').get()
        
        #     yield {
        #         'title': title,
        #         'content': content
        #     }