from scrapy import signals
from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings
from ArkEnergyScraper.Scraper.spiders.ArkSpiders import Spider1
from ArkEnergyScraper.Scraper.spiders.ArkSpiders import Spider2
from pptx import Presentation
from pptx.util import Pt
from pptx.util import Inches
import os

def slide1(prs):
    # Set background image
    image_path = "background_image.jpg"  # Path to your image file
    left = top = Inches(0)
    width = height = Inches(10)
    slide.shapes.add_picture(image_path, left, top, width=width, height=height)

    # Slide 1: Title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "ARK Energy"

    # Get the current directory of the Python script
    current_dir = os.path.dirname(os.path.abspath(__file__))
    img_folder = "Logo"

    # Path to the image file relative to the current directory
    img_filename1 = "180degreelogo.png"
    img_path1 = os.path.join(current_dir, img_folder, img_filename1)
    from_left = Inches(7.3)
    from_top = Inches(6.5)
    length = Inches(2)
    slide.shapes.add_picture(img_path1, from_left, from_top, length)

    img_filename2 = "Arkenergylogo.png"
    img_path2 = os.path.join(current_dir, img_folder, img_filename2)
    from_left = Inches(5)
    from_top = Inches(6.5)
    length = Inches(2)
    slide.shapes.add_picture(img_path2, from_left, from_top, length)

def slide3(prs):
    # Add a slide to the presentation
    slide_layout = prs.slide_layouts[5]  # Choose a layout that supports tables (e.g., Title and Content)
    slide = prs.slides.add_slide(slide_layout)

    # Define table dimensions and position
    rows = 2
    cols = 3
    left = Inches(1)
    top = Inches(2)
    width = Inches(6)
    height = Inches(1.5)

    # Add a table to the slide
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Populate the table with data
    table.cell(0, 0).text = "Header 1"
    table.cell(0, 1).text = "Header 2"
    table.cell(0, 2).text = "Header 3"

    table.cell(1, 0).text = "Data 1"
    table.cell(1, 1).text = "Data 2"
    table.cell(1, 2).text = "Data 3"

def generate_ppt(slides_data):
    prs = Presentation()
    slide1(prs)
    
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
    
    slide3(prs)
    prs.save('Monthly_Report.pptx')

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