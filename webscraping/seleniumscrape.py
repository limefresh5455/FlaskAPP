from selenium import webdriver

# Set up the WebDriver (this example uses Chrome)
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Run in headless mode
driver = webdriver.Chrome(options=options)

# Open the web page
url = 'https://midone-vue.vercel.app/'
driver.get(url)

# Wait for the page to fully render
driver.implicitly_wait(10)

# Get the page source (including dynamically generated content)
page_html = driver.page_source

# Save the HTML to a file
with open('page.html', 'w', encoding='utf-8') as f:
    f.write(page_html)

# Close the WebDriver
driver.quit()
