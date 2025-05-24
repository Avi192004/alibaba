from datetime import datetime, timedelta
import subprocess
import sys
import os
import json
import time
import random
import traceback
import psutil
import requests
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import openpyxl
from openpyxl import Workbook

# ------------------ AUTO-INSTALL REQUIRED MODULES ------------------

REQUIRED_PACKAGES = [
    "undetected-chromedriver",
    "selenium",
    "psutil",
    "requests"
]

for package in REQUIRED_PACKAGES:
    try:
        __import__(package.replace("-", "_"))
    except ImportError:
        print(f"📦 Installing missing package: {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

import undetected_chromedriver as uc

print("✅ All required packages are installed!")

# ------------------ CONFIGURATION ------------------

MAIN_URL = "https://onetalk.alibaba.com/message/weblitePWA.htm?isGray=1&from=menu&hideMenu=1#/"
BASE_URL = "https://alibaba.com/"
RAG_URL = "https://609f-34-59-106-222.ngrok-free.app/search-embed"  # Replace with your real endpoint
USE_AI = True  # Toggle AI replies

REPLIES = [
    "Hello! Thanks for your inquiry. Our team will assist you shortly.",
    "Hi there! Your inquiry is important to us. We'll be with you shortly.",
    "Greetings! Thank you for reaching out. One of our representatives will assist you soon.",
    "Hey! Thanks for getting in touch. We'll be happy to help you shortly.",
    "Hi! We appreciate your message. Our team will assist you as soon as possible.",
    "Hello! Thanks for your inquiry. Please hold on, our team will assist you soon."
]

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

COOKIES_FILE = os.path.join(BASE_DIR, "cookies.json")
ERROR_LOG = os.path.join(BASE_DIR, "error.log")
ACTIVITY_LOG = os.path.join(BASE_DIR, "activity.log")
# Spreadsheet setup
SHEET_FILE = os.path.join(BASE_DIR, "inquiries.xlsx")

CHROME_PID = None

# -------------------- EXCEL SETUP ------------------

if not os.path.exists(SHEET_FILE):
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Inquiries"
    sheet.append([
        "Inquiry ID", "User", "Country", "Company", "Email", "Registration Date",
        "Product Views", "Inquiries", "RFQs", "Login Days",
        "Spam Inquiries", "Blacklist Count", "Follow-up Date", "Count"
    ])
    wb.save(SHEET_FILE)
else:
    wb = openpyxl.load_workbook(SHEET_FILE)
    sheet = wb.active

# ------------------ LOGGING ------------------

def log_error(error_message):
    with open(ERROR_LOG, "a", encoding="utf-8") as log_file:
        log_file.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - ERROR: {error_message}\n")
        log_file.write(traceback.format_exc() + "\n\n")
    print(f"[❌ ERROR] {error_message}")

def log_activity(message):
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
    with open(ACTIVITY_LOG, "a", encoding="utf-8") as log:
        log.write(f"{timestamp} - {message}\n")
    print(f"📘 {message}")

def wait_for_user_confirmation(message):
    print(f"[ℹ️] {message}")
    input("🔄 Press Enter once done...")

def cleanup_and_exit():
    global CHROME_PID
    if CHROME_PID:
        try:
            chrome_process = psutil.Process(CHROME_PID)
            chrome_process.terminate()
            chrome_process.wait(timeout=5)
            log_activity("✅ Closed Chrome process started by script.")
        except psutil.NoSuchProcess:
            log_activity("⚠️ Chrome process not found (maybe already exited).")
        except Exception as e:
            log_error(f"⚠️ Error closing Chrome: {str(e)}")
    sys.exit(1)

# ------------------ BROWSER SETUP ------------------

def start_browser():
    global CHROME_PID
    try:
        options = uc.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-extensions")
        options.add_argument("--headless=new")

        driver = uc.Chrome(options=options)
        CHROME_PID = driver.browser_pid
        log_activity(f"🔵 Started Chrome with PID: {CHROME_PID}")
        return driver
    except Exception as e:
        log_error(f"⚠️ Failed to start browser: {str(e)}")
        cleanup_and_exit()

# ------------------ LOGIN ------------------

def login(driver):
    try:
        driver.get(BASE_URL)
        if os.path.exists(COOKIES_FILE):
            with open(COOKIES_FILE, "r") as f:
                cookies = json.load(f)
                for cookie in cookies:
                    driver.add_cookie(cookie)
            log_activity("✅ Cookies loaded.")
        else:
            wait_for_user_confirmation("🔐 No cookies found. Please log in manually in the browser window.")
            driver.get(MAIN_URL)
            time.sleep(10)
            cookies = driver.get_cookies()
            with open(COOKIES_FILE, "w") as f:
                json.dump(cookies, f)
            log_activity("✅ Cookies saved after manual login.")
        driver.get(MAIN_URL)
    except Exception as e:
        log_error(f"⚠️ Login failed: {str(e)}")
        cleanup_and_exit()

# ------------------ API RESPONSE ------------------

def get_api_response(question, img_url=None):
    try:
        url = RAG_URL
        headers = {"Content-Type": "application/json"}
        payload = {"query": question}
        if img_url:
            payload["image"] = img_url

        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()

        data = response.json()
        message = data.get("answer", "We’ll get back to you shortly.")
        log_activity(f"🔍 API response: {message[:60]}...")
        return message
    except Exception as e:
        log_error(f"❌ API request failed: {str(e)}")
        return None

def get_ai_response(driver):
    try:
        ai_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "assistant-entry-icon"))
        )
        ai_button.click()
        log_activity("🤖 Clicked AI Assistant.")

        use_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Use this')]"))
        )
        use_btn.click()
        log_activity("✅ Inserted AI-generated message.")

        time.sleep(2)
        pre = driver.find_element(By.CSS_SELECTOR, "#send-box-wrapper pre")
        ai_text = pre.get_attribute("textContent").strip()
        log_activity(f"🤖 AI reply preview: {ai_text[:60]}...")
        return ai_text
    except Exception as e:
        log_activity("⚠️ AI Assistant fallback failed.")
        return None

def generate_reply(driver, query, img_url):
    reply = None
    if USE_AI:
        reply = get_api_response(query, img_url)
    if not reply or reply.strip() == "":
        reply = get_ai_response(driver)
    if not reply or reply.strip() == "":
        reply = random.choice(REPLIES)
    return reply

def send_message(driver, recipient, message):
    try:
        message_box = driver.find_element(By.CLASS_NAME, "send-textarea")
        message_box.send_keys(Keys.CONTROL + "a")
        time.sleep(0.5)
        message_box.send_keys(Keys.BACKSPACE)
        time.sleep(2)
        message_box.send_keys(message)
        time.sleep(random.uniform(1, 3))
        send_button = driver.find_element(By.XPATH, "//button[contains(@class, 'send-tool-button')]")
        send_button.click()
        log_activity(f"✅ Sent message to {recipient}: {message}")
    except Exception as e:
        log_error(f"❌ Error sending message to {recipient}: {str(e)}")

def extract_message_data(message_container):
    try:
        msg_type = json.loads(message_container.get_attribute("data-expinfo"))['messageType']
        message_text = ''
        image_url = None

        if msg_type == 1:
            message_text = message_container.find_element(By.CLASS_NAME, 'session-rich-content').text
        elif msg_type in [2000, 60]:
            image_element = message_container.find_element(By.XPATH, "//div[@view-name='ImageView']/div/img | //div/img")
            image_url = image_element.get_attribute("src")
            message_text = "details on this product"
        elif msg_type in [50, 63]:
            message_text = message_container.find_element(By.CLASS_NAME, 'description-container').text
            image_element = message_container.find_element(By.XPATH, '//p/img')
            image_url = image_element.get_attribute("src")
        elif msg_type == 61:
            file_details = json.loads(message_container.find_element(By.XPATH, '//div[@data-exp="card-file"]').get_attribute("data-query"))
            message_text = f"File: {file_details.get('fileName')} ({file_details.get('fileSize')})"
        elif msg_type == 57:
            message_text = ''  # Skip business cards

        return message_text, image_url
    except Exception as e:
        log_error(f"❌ Error extracting message data: {str(e)}")
        return None, None

def store_inquiry(driver):
    try:
        user = driver.find_element(By.CSS_SELECTOR, ".name-text").text.strip()
        country = driver.find_element(By.CSS_SELECTOR, ".country-flag-label").text.strip()
        info_array = driver.find_elements(By.CSS_SELECTOR, "div.base-information-form-item-content > span")
        company = info_array[0].text.strip() if len(info_array) > 0 else ""
        email = info_array[1].text.strip() if len(info_array) > 1 else ""
        registration_date = info_array[2].text.strip() if len(info_array) > 2 else ""

        product_views_count = driver.find_element(By.CSS_SELECTOR, "div.product-visit.indicator > div.count").text.strip()
        inquiries_count = driver.find_element(By.CSS_SELECTOR, "div.inquiries-count.indicator > div.count").text.strip()
        available_rfq_count = driver.find_element(By.CSS_SELECTOR, "div.availble-rfq.indicator > div.count").text.strip()
        login_days_count = driver.find_element(By.CSS_SELECTOR, "div.landing-days.indicator > div.count").text.strip()
        spam_inquiries_count = driver.find_element(By.CSS_SELECTOR, "div.trash-inquires.indicator > div.count").text.strip()
        blacklist_count = driver.find_element(By.CSS_SELECTOR, "div.add-blacklist.indicator > div.count").text.strip()

        follow_up_date = (datetime.today() + timedelta(days=3)).strftime('%Y-%m-%d')
        inquiry_id = f"INQ-{int(time.time())}"
        count = 1

        # Save to Excel sheet
        sheet.append([
            inquiry_id, user, country, company, email, registration_date,
            product_views_count, inquiries_count, available_rfq_count,
            login_days_count, spam_inquiries_count, blacklist_count,
            follow_up_date, count
        ])
        wb.save(SHEET_FILE)
        log_activity(f"📝 Stored inquiry {inquiry_id} to sheet.")

        # Send to n8n webhook
        webhook_url = "https://n8n.ecowoodies.com/webhook/alibabadumping"  # Replace with actual URL
        payload = {
            "inquiry_id": inquiry_id,
            "user": user,
            "country": country,
            "company": company,
            "email": email,
            "registration_date": registration_date,
            "product_views_count": product_views_count,
            "inquiries_count": inquiries_count,
            "available_rfq_count": available_rfq_count,
            "login_days_count": login_days_count,
            "spam_inquiries_count": spam_inquiries_count,
            "blacklist_count": blacklist_count,
            "follow_up_date": follow_up_date,
            "count": count
        }

        response = requests.post(webhook_url, json=payload)
        if response.status_code == 200:
            log_activity(f"📡 Inquiry {inquiry_id} sent to webhook.")
        else:
            log_error(f"❌ Failed to send inquiry to webhook: {response.text}")

    except Exception as e:
        log_error(f"❌ Error storing/sending inquiry: {str(e)}")


# ------------------ MAIN LOOP ------------------

def main():
    driver = start_browser()
    if not driver:
        print("❌ Failed to start browser.")
        return

    login(driver)

    close_pop = driver.find_elements(By.CLASS_NAME, "im-next-dialog-close")
    if close_pop:
        close_pop[0].click()
        log_activity("🔒 Closed pop-up.")

    close_pop = driver.find_elements(By.CLASS_NAME, "close-icon")
    if close_pop:
        close_pop[0].click()
        log_activity("🔒 Closed pop-up.")

    i = 0
    while True:
        try:
            unread_messages = driver.find_elements(By.CLASS_NAME, "unread-num")
            unread_messages_without_labels = []

            for message in unread_messages:
                is_inquiry=False
                container = message.find_element(By.XPATH, "ancestor::div[2]")
                if container.find_element(By.CLASS_NAME, 'latest-msg-oneline').text == "[Inquiry]":
                    is_inquiry=True
                
                labels = container.find_elements(By.CLASS_NAME, "tag-item ")
                last_msg_time = container.find_element(By.CLASS_NAME, "contact-time").text
                today = datetime.today()
                msg_dt = datetime.strptime(last_msg_time, "%H:%M").replace(year=today.year, month=today.month, day=today.day)
                msg_timestamp = msg_dt.timestamp()
                if not labels or time.time() - 180 > msg_timestamp:
                    unread_messages_without_labels.append(message)

            if unread_messages_without_labels:
                i = 0
                container = unread_messages_without_labels[0].find_element(By.XPATH, "ancestor::div[2]")
                recipient = container.get_attribute("data-name")
                log_activity(f"📨 New unread message from: {recipient}")
                unread_messages_without_labels[0].click()
                time.sleep(random.uniform(2, 5))

                message_container = driver.find_element(By.CSS_SELECTOR, "div.scroll-box > *")
                message_text, img_url = extract_message_data(message_container)
                reply = generate_reply(driver, message_text, img_url)
                send_message(driver, recipient, reply)
                if is_inquiry:
                    log_activity("🔄 Inquiry detected, refreshing main page.")
                    store_inquiry(driver)

                driver.get(MAIN_URL)
                log_activity("🔄 Returned to main page.")

            time.sleep(random.uniform(10, 15))
            i += 1
            if i > 7:
                log_activity("🔄 Refreshing main page after inactivity.")
                driver.refresh()
                time.sleep(random.uniform(25, 30))
        except Exception as e:
            log_error(f"⚠️ Critical error in main loop: {str(e)}")
            cleanup_and_exit()
            break

if __name__ == "__main__":
    main()
