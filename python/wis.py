import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), ".venv/Lib/site-packages"))

import chromedriver_binary
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By

ID = "1368040365050" # ohmachi dam
DT1 = "20220819"
DT2 = "20220825"

load_url = "http://www1.river.go.jp/cgi-bin/DspDamData.exe?KIND=1" + \
    "&ID=" + ID + "&BGNDATE=" + DT1 + "&ENDDATE=" + DT2
# http://www1.river.go.jp/cgi-bin/DspDamData.exe?KIND=1&ID=1368040365050&BGNDATE=20220819&ENDDATE=20220825

options = webdriver.ChromeOptions()
options.add_argument("--headless")
driver = webdriver.Chrome(options=options)
driver.get(load_url)

if 1 == 2: # parse anchor element

    element = driver.find_element(By.TAG_NAME, "a")
    element.click()

    handle_array = driver.window_handles
    driver.switch_to.window(handle_array[-1]) # focus newly opening tab

    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')

    contents = soup.find("pre").contents[0]
    with open("contents.csv", "w", encoding="utf-8") as f:
        print(contents, file=f)

else: # parse iframe element

    element = driver.find_element(By.TAG_NAME, "iframe")
    driver.switch_to.frame(element)
    html = driver.page_source.encode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    for tr in soup.find_all("tr"):
        cols = []
        for td in tr.find_all("td"):
            try:
                contents = td.find("font").contents[0]
            except:
                contents = td.contents[0]
            cols.append(contents)
        rows.append(cols)

    import csv
    with open("contents.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(rows)

driver.quit()