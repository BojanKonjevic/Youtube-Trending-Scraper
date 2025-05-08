import time
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager

YOUTUBE_URL = "https://youtube.com/"
EXCEL_FILENAME = "Youtube Trending.xlsx"
JSON_FILENAME = "Youtube Trending.json"

def setup_driver(headless=True):
    options = Options()
    options.add_argument("--start-maximized")
    if headless:
        options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def go_to_trending(driver):
    driver.get(YOUTUBE_URL)
    trending_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[@title='Trending']"))
    )
    trending_button.click()
    WebDriverWait(driver, 10).until(EC.url_changes(YOUTUBE_URL))

def scroll_to_load_all(driver, pause=0.3, step=500):
    current_position = 0
    while True:
        driver.execute_script(f"window.scrollTo(0, {current_position});")
        current_position += step
        time.sleep(pause)
        new_height = driver.execute_script("return document.documentElement.scrollHeight")
        if current_position >= new_height:
            break

def extract_videos(driver):
    video_divs = driver.find_elements(By.XPATH, "//div[@id='dismissible' and @class='style-scope ytd-video-renderer']")
    videos = []

    for div in video_divs:
        try:
            title_element = div.find_element(By.XPATH, ".//a[@id='video-title']")
            channel_element = div.find_element(By.ID, 'channel-name').find_element(By.TAG_NAME, 'a')
            metadata = div.find_element(By.ID, 'metadata-line').find_elements(By.CSS_SELECTOR, ".inline-metadata-item")

            video_data = {
                'title': title_element.get_attribute('title'),
                'channel_name': channel_element.text,
                'viewcount': metadata[0].text if len(metadata) > 0 else '',
                'date': metadata[1].text if len(metadata) > 1 else '',
                'link': title_element.get_attribute('href'),
                'thumbnail': div.find_element(By.TAG_NAME, 'img').get_attribute('src')
            }
            videos.append(video_data)
        except Exception as e:
            print(f"Error extracting video: {e}")

    return videos

def save_to_excel(videos, filename):
    df = pd.DataFrame(videos)
    df.to_excel(filename, index=False)

    wb = load_workbook(filename)
    ws = wb.active

    for column_cells in ws.columns:
        max_length = max((len(str(cell.value)) for cell in column_cells), default=0)
        adjusted_width = max_length + 6
        column_letter = column_cells[0].column_letter
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(filename)

def save_to_json(videos, filename):
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(videos, f, ensure_ascii=False, indent=4)

def print_videos(videos):
    for video in videos:
        print(f"Title: {video['title']}")
        print(f"Channel: {video['channel_name']}")
        print(f"View Count: {video['viewcount']}")
        print(f"Date: {video['date']}")
        print(f"Link: {video['link']}")
        print(f"Thumbnail: {video['thumbnail']}")
        print("-" * 150)

def main():
    driver = setup_driver(headless=True)
    try:
        go_to_trending(driver)
        scroll_to_load_all(driver)
        videos = extract_videos(driver)

        if not videos:
            print("No videos extracted.")
            return

        save_to_excel(videos, EXCEL_FILENAME)
        save_to_json(videos, JSON_FILENAME)
        print_videos(videos)

        print(f"\nSaved {len(videos)} videos to '{EXCEL_FILENAME}' and '{JSON_FILENAME}'")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
