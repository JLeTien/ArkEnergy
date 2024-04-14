from scrapy import signals
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings
from Scraper.ArkEnergyScraper.spiders.ArkSpiders import Spider1
from Scraper.ArkEnergyScraper.spiders.ArkSpiders import Spider2
from pptx import Presentation
from pptx.util import Pt

def generate_ppt(slides_data):
    prs = Presentation()
    
    for slide_content in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        title, content = slide_content['title'], slide_content['content']
        title_shape = slide.shapes.title
        title_shape.text = title
        title_text_frame = title_shape.text_frame
        title_paragraph = title_text_frame.paragraphs[0]
        title_font = title_paragraph.font
        title_font.size = Pt(12)  # Set the font size (e.g., 24 points)

        content_shape = slide.placeholders[1]
        content_shape.text = content
        content_text_frame = content_shape.text_frame
        content_paragraph = content_text_frame.paragraphs[0]
        content_font = content_paragraph.font
        content_font.size = Pt(10)  # Set the font size (e.g., 18 points)
    
    prs.save('example.pptx')

def main():
    process = CrawlerProcess()
    process.crawl(Spider1)
    process.start()
    
    # After the crawling process has finished, you can access the slides_data attribute
    slides_data = Spider1.slides_data
    
    # Now you can use the slides_data list as needed
    print(slides_data)
    generate_ppt(slides_data)
    
if __name__ == "__main__":
    main()