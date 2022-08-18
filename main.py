from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from time import sleep
import xlsxwriter


PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

driver.get("https://www.youtube.com/")
driver.maximize_window()
search = driver.find_element(By.NAME, "search_query")

all_channels = [{
    "Channel_name": "DNF",
    "Video_title": "DNF",
    "Views": "",
    "Publish_date": "DNF"
} for video in range(100)]


print("--------------")
search_key = input("What do you want to search? \n--------------\n")

search.send_keys(search_key)
sleep(1)
search.send_keys(Keys.RETURN)
sleep(3)
main = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "content"))
)

sleep(1)

driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

sleep(3)

videos = main.find_elements(
    By.CLASS_NAME, "style-scope ytd-item-section-renderer")


print("------------------------------------------")

videodivs = videos[0].find_elements(By.ID, "dismissible")
vid_counter = 0

for video in videodivs:
    try:
        title = video.find_element(By.ID, "video-title")
        channel = video.find_element(
            By.ID, "text-container").find_element(By.ID, "text").find_element(By.TAG_NAME, "a").get_attribute('innerHTML')
        data = video.find_element(By.ID, "metadata-line").text
        views, published = data.split("\n")
        all_channels[vid_counter] = {
            "Channel_name": f"{channel}",
            "Video_title": f"{title.text}",
            "Views": f"{views}",
            "Publish_date": f"{published}"
        }
    except:
        continue
    vid_counter += 1
    print(vid_counter, channel)
    print("--------------")


sleep(1)


driver.quit()


def generate_excel(workbook_name: str, worksheet_name: str, headers_list: list, data: list):
    # Creating workbook
    workbook = xlsxwriter.Workbook(workbook_name)
    # Creating worksheet
    worksheet = workbook.add_worksheet(worksheet_name)
    # Adding data
    for index1, entry in enumerate(data):
        for index2, header in enumerate(headers_list):
            worksheet.write(index1+1, index2, entry[header])
    # Fixing Headers
    worksheet.write(0, 0, headers_list[0])
    worksheet.write(0, 1, headers_list[1])
    worksheet.write(0, 2, headers_list[2])
    worksheet.write(0, 3, headers_list[3])
    # Close workbook
    workbook.close()


generate_excel(f"{search_key.replace(' ','_')}.xlsx", "firstsheet",
               ["Channel_name", "Video_title", "Views", "Publish_date"], all_channels)
