import scrapy

class MySpider(scrapy.Spider):
    name = 'myspider'
    allowed_domains = ['kreyonsystem.com']
    start_urls = ['https://kreyonsystems.com/']

    def parse(self, response):
        # Extract the HTML content
        page_html = response.text

        # Debugging: Print a message to confirm the method is called
        print("Parsing the page and saving HTML")

        # Save the HTML to a file
        with open('kreyon.html', 'w', encoding='utf-8') as f:
            f.write(page_html)

        # Optionally, you can also yield the HTML content as an item
        yield {'html': page_html}
