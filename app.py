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
from selenium.common.exceptions import (
    NoSuchElementException, 
    TimeoutException, 
    StaleElementReferenceException,
    InvalidSessionIdException,
    WebDriverException
)
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

CHROME_PID = None
MAX_SESSION_RECOVERY_ATTEMPTS = 3
SESSION_CHECK_INTERVAL = 300  # Check session health every 5 minutes

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
    cleanup_our_chrome_process()
    sys.exit(1)

# ------------------ SESSION MANAGEMENT ------------------

def is_session_valid(driver):
    """Check if the current session is still valid"""
    try:
        # Try a simple operation to test session validity
        driver.current_url
        driver.title
        return True
    except (InvalidSessionIdException, WebDriverException):
        return False
    except Exception as e:
        log_activity(f"⚠️ Unexpected error checking session: {str(e)}")
        return False

def cleanup_our_chrome_process():
    """Only kill Chrome processes that were started by this script"""
    global CHROME_PID
    if CHROME_PID:
        try:
            chrome_process = psutil.Process(CHROME_PID)
            if chrome_process.is_running():
                chrome_process.terminate()
                chrome_process.wait(timeout=5)
                log_activity(f"🧹 Cleaned up our Chrome process (PID: {CHROME_PID})")
        except (psutil.NoSuchProcess, psutil.TimeoutExpired):
            log_activity(f"⚠️ Chrome process {CHROME_PID} already terminated")
        except Exception as e:
            log_activity(f"⚠️ Error cleaning up Chrome process {CHROME_PID}: {str(e)}")
        finally:
            CHROME_PID = None

def start_browser():
    """Start browser with enhanced error handling"""
    global CHROME_PID
    
    # Only clean up our own Chrome process if it exists
    cleanup_our_chrome_process()
    
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            options = uc.ChromeOptions()
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-web-security")
            options.add_argument("--allow-running-insecure-content")
            options.add_argument("--disable-features=VizDisplayCompositor")
            options.add_argument("--headless=new")

            driver = uc.Chrome(options=options, version_main=None)
            CHROME_PID = driver.browser_pid
            log_activity(f"🔵 Started Chrome with PID: {CHROME_PID} (attempt {attempt + 1})")
            
            # Test the session immediately
            if is_session_valid(driver):
                return driver
            else:
                driver.quit()
                continue
                
        except Exception as e:
            log_error(f"⚠️ Failed to start browser (attempt {attempt + 1}): {str(e)}")
            if attempt < max_attempts - 1:
                time.sleep(5)  # Wait before retrying
                continue
            else:
                cleanup_and_exit()
    
    return None

def recover_session(driver):
    """Attempt to recover from a broken session"""
    global CHROME_PID
    
    log_activity("🔄 Attempting session recovery...")
    
    try:
        # Try to quit the current driver gracefully
        if driver:
            try:
                driver.quit()
            except:
                pass
    except:
        pass
    
    # Only kill our own Chrome process
    cleanup_our_chrome_process()
    
    # Wait a bit for cleanup
    time.sleep(3)
    
    # Start a new browser session
    new_driver = start_browser()
    if new_driver:
        log_activity("✅ Session recovered successfully")
        return new_driver
    else:
        log_error("❌ Failed to recover session")
        cleanup_and_exit()

# ------------------ LOGIN ------------------

def login(driver):
    """Login with enhanced error handling"""
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            if not is_session_valid(driver):
                log_activity("⚠️ Invalid session detected during login")
                return False
                
            driver.get(BASE_URL)
            time.sleep(3)
            
            if os.path.exists(COOKIES_FILE):
                with open(COOKIES_FILE, "r") as f:
                    cookies = json.load(f)
                    for cookie in cookies:
                        try:
                            driver.add_cookie(cookie)
                        except Exception as e:
                            # Skip invalid cookies
                            continue
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
            time.sleep(5)
            
            # Verify we're logged in by checking for expected elements
            if is_session_valid(driver):
                return True
            else:
                if attempt < max_attempts - 1:
                    log_activity(f"⚠️ Login verification failed, retrying... (attempt {attempt + 1})")
                    time.sleep(5)
                    continue
                    
        except Exception as e:
            log_error(f"⚠️ Login failed (attempt {attempt + 1}): {str(e)}")
            if attempt < max_attempts - 1:
                time.sleep(5)
                continue
            else:
                return False
    
    return False

# ------------------ API RESPONSE ------------------

def get_api_response(question, img_url=None):
    try:
        url = RAG_URL
        headers = {"Content-Type": "application/json"}
        payload = {"query": question}
        if img_url:
            payload["image"] = img_url

        response = requests.post(url, headers=headers, json=payload, timeout=10)
        response.raise_for_status()

        data = response.json()
        message = data.get("answer", "We'll get back to you shortly.")
        log_activity(f"🔍 API response: {message[:60]}...")
        return message
    except Exception as e:
        log_error(f"❌ API request failed: {str(e)}")
        return None

def get_ai_response(driver):
    try:
        if not is_session_valid(driver):
            return None
            
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
    # if not reply or reply.strip() == "":
    #     reply = get_ai_response(driver)
    if not reply or reply.strip() == "":
        reply = random.choice(REPLIES)
    return reply

def send_message(driver, recipient, message):
    try:
        if not is_session_valid(driver):
            return False
            
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
        return True
    except Exception as e:
        log_error(f"❌ Error sending message to {recipient}: {str(e)}")
        return False

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

def safe_find_element(element, by, value, default=""):
    """Safely find an element and return its text, or default value if not found"""
    try:
        return element.find_element(by, value).text.strip()
    except (NoSuchElementException, StaleElementReferenceException, InvalidSessionIdException):
        return default

def safe_find_elements(driver_or_element, by, value):
    """Safely find elements and return the list, or empty list if not found"""
    try:
        if not is_session_valid(driver_or_element) if hasattr(driver_or_element, 'current_url') else True:
            return []
        return driver_or_element.find_elements(by, value)
    except (NoSuchElementException, StaleElementReferenceException, InvalidSessionIdException):
        return []

def check_if_inquiry(container):
    """Check if message is an inquiry with proper error handling"""
    try:
        # Try multiple possible selectors for the message type
        selectors_to_try = [
            'latest-msg-oneline',
            'latest-msg',
            'msg-content',
            'message-content',
            'session-content'
        ]
        
        message_text = ""
        for selector in selectors_to_try:
            try:
                element = container.find_element(By.CLASS_NAME, selector)
                message_text = element.text.strip()
                break
            except NoSuchElementException:
                continue
        
        # Check if it's an inquiry based on text content
        inquiry_keywords = ["[Inquiry]", "[Product]", "inquiry", "product"]
        return any(keyword.lower() in message_text.lower() for keyword in inquiry_keywords)
        
    except Exception as e:
        log_activity(f"⚠️ Could not determine if message is inquiry: {str(e)}")
        return False

def store_inquiry(driver, img_url):
    try:
        if not is_session_valid(driver):
            return False
            
        user = safe_find_element(driver, By.CSS_SELECTOR, ".name-text", "Unknown User")
        country = safe_find_element(driver, By.CSS_SELECTOR, ".country-flag-label", "Unknown Country")
        
        info_array = safe_find_elements(driver, By.CSS_SELECTOR, "div.base-information-form-item-content > span")
        company = info_array[0].text.strip() if len(info_array) > 0 else ""
        email = info_array[1].text.strip() if len(info_array) > 1 else ""
        registration_date = info_array[2].text.strip() if len(info_array) > 2 else ""

        product_views_count = safe_find_element(driver, By.CSS_SELECTOR, "div.product-visit.indicator > div.count", "0")
        inquiries_count = safe_find_element(driver, By.CSS_SELECTOR, "div.inquiries-count.indicator > div.count", "0")
        available_rfq_count = safe_find_element(driver, By.CSS_SELECTOR, "div.availble-rfq.indicator > div.count", "0")
        login_days_count = safe_find_element(driver, By.CSS_SELECTOR, "div.landing-days.indicator > div.count", "0")
        spam_inquiries_count = safe_find_element(driver, By.CSS_SELECTOR, "div.trash-inquires.indicator > div.count", "0")
        blacklist_count = safe_find_element(driver, By.CSS_SELECTOR, "div.add-blacklist.indicator > div.count", "0")
        
        follow_up_date = (datetime.today() + timedelta(days=3)).strftime('%Y-%m-%d')
        inquiry_id = f"INQ-{int(time.time())}"
        count = 1

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
            "count": count,
            "img": img_url
        }

        try:
            response = requests.post(webhook_url, json=payload, timeout=10)
            if response.status_code == 200:
                log_activity(f"📡 Inquiry {inquiry_id} sent to webhook.")
            else:
                log_error(f"❌ Failed to send inquiry to webhook: {response.text}")
        except requests.exceptions.RequestException as e:
            log_error(f"❌ Webhook request failed: {str(e)}")
        
        return True

    except Exception as e:
        log_error(f"❌ Error storing/sending inquiry: {str(e)}")
        return False

# ------------------ MAIN LOOP ------------------

def main():
    driver = start_browser()
    if not driver:
        print("❌ Failed to start browser.")
        return

    if not login(driver):
        log_error("❌ Failed to login")
        cleanup_and_exit()

    # Close popups with error handling
    try:
        close_pop = safe_find_elements(driver, By.CLASS_NAME, "im-next-dialog-close")
        if close_pop:
            close_pop[0].click()
            log_activity("🔒 Closed pop-up.")
    except Exception as e:
        log_activity(f"⚠️ Could not close first popup: {str(e)}")

    try:
        close_pop = safe_find_elements(driver, By.CLASS_NAME, "close-icon")
        if close_pop:
            close_pop[0].click()
            log_activity("🔒 Closed pop-up.")
    except Exception as e:
        log_activity(f"⚠️ Could not close second popup: {str(e)}")

    i = 0
    consecutive_errors = 0
    last_session_check = time.time()
    session_recovery_attempts = 0
    
    while True:
        try:
            # Periodic session health check
            current_time = time.time()
            if current_time - last_session_check > SESSION_CHECK_INTERVAL:
                if not is_session_valid(driver):
                    log_activity("⚠️ Session health check failed")
                    raise InvalidSessionIdException("Session invalid during health check")
                last_session_check = current_time
                log_activity("✅ Session health check passed")

            unread_messages = safe_find_elements(driver, By.CLASS_NAME, "unread-num")
            unread_messages_without_labels = []

            for message in unread_messages:
                try:
                    is_inquiry = False
                    container = message.find_element(By.XPATH, "ancestor::div[2]")
                    
                    # Check if it's an inquiry with error handling
                    is_inquiry = check_if_inquiry(container)
                    
                    labels = safe_find_elements(container, By.CLASS_NAME, "tag-item")
                    
                    # Get message time with error handling
                    try:
                        last_msg_time = safe_find_element(container, By.CLASS_NAME, "contact-time")
                        if last_msg_time:
                            today = datetime.today()
                            msg_dt = datetime.strptime(last_msg_time, "%H:%M").replace(year=today.year, month=today.month, day=today.day)
                            msg_timestamp = msg_dt.timestamp()
                            
                            # Only process if no labels or message is recent
                            if not labels or time.time() - 180 > msg_timestamp:
                                unread_messages_without_labels.append((message, is_inquiry))
                        else:
                            # If we can't get time, still process the message
                            if not labels:
                                unread_messages_without_labels.append((message, is_inquiry))
                    except (ValueError, AttributeError) as e:
                        log_activity(f"⚠️ Could not parse message time: {str(e)}")
                        # If we can't parse time, still process if no labels
                        if not labels:
                            unread_messages_without_labels.append((message, is_inquiry))
                            
                except (NoSuchElementException, StaleElementReferenceException) as e:
                    log_activity(f"⚠️ Stale element encountered, skipping message: {str(e)}")
                    continue

            if unread_messages_without_labels:
                i = 0
                consecutive_errors = 0  # Reset error counter on success
                session_recovery_attempts = 0  # Reset recovery attempts
                
                message_element, is_inquiry = unread_messages_without_labels[0]
                try:
                    container = message_element.find_element(By.XPATH, "ancestor::div[2]")
                    recipient = container.get_attribute("data-name") or "Unknown Recipient"
                    log_activity(f"📨 New unread message from: {recipient}")
                    
                    message_element.click()
                    time.sleep(random.uniform(2, 5))

                    # Try to extract message data
                    try:
                        message_container = driver.find_element(By.CSS_SELECTOR, "div.scroll-box > *")
                        message_text, img_url = extract_message_data(message_container)
                    except NoSuchElementException:
                        message_text, img_url = "New message", None
                        log_activity("⚠️ Could not extract message details, using default.")

                    reply = generate_reply(driver, message_text, img_url)
                    if send_message(driver, recipient, reply):
                        if is_inquiry:
                            log_activity("🔄 Inquiry detected, storing data.")
                            store_inquiry(driver, img_url)

                    driver.get(MAIN_URL)
                    log_activity("🔄 Returned to main page.")
                    
                except (NoSuchElementException, StaleElementReferenceException) as e:
                    log_activity(f"⚠️ Element became stale, refreshing page: {str(e)}")
                    driver.get(MAIN_URL)
                    time.sleep(5)

            time.sleep(random.uniform(10, 15))
            i += 1
            
            # Refresh page periodically
            if i > 7:
                log_activity("🔄 Refreshing main page after inactivity.")
                if is_session_valid(driver):
                    driver.refresh()
                    time.sleep(random.uniform(25, 30))
                else:
                    raise InvalidSessionIdException("Session invalid during refresh")
                i = 0
                
        except InvalidSessionIdException as e:
            log_error(f"⚠️ Session disconnected: {str(e)}")
            if session_recovery_attempts < MAX_SESSION_RECOVERY_ATTEMPTS:
                session_recovery_attempts += 1
                driver = recover_session(driver)
                if driver and login(driver):
                    consecutive_errors = 0
                    continue
                else:
                    log_error("❌ Failed to recover session and login")
                    cleanup_and_exit()
            else:
                log_error("❌ Maximum session recovery attempts reached")
                cleanup_and_exit()
                
        except Exception as e:
            consecutive_errors += 1
            log_error(f"⚠️ Error in main loop (#{consecutive_errors}): {str(e)}")
            
            # If too many consecutive errors, try to recover
            if consecutive_errors >= 5:
                log_activity("🔄 Too many consecutive errors, attempting recovery...")
                try:
                    if is_session_valid(driver):
                        driver.get(MAIN_URL)
                        time.sleep(10)
                        consecutive_errors = 0
                    else:
                        # Session is invalid, attempt recovery
                        if session_recovery_attempts < MAX_SESSION_RECOVERY_ATTEMPTS:
                            session_recovery_attempts += 1
                            driver = recover_session(driver)
                            if driver and login(driver):
                                consecutive_errors = 0
                                continue
                        raise Exception("Session recovery failed")
                except Exception as recovery_error:
                    log_error(f"❌ Recovery failed: {str(recovery_error)}")
                    cleanup_and_exit()
            else:
                # Wait a bit before retrying
                time.sleep(random.uniform(30, 60))

if __name__ == "__main__":
    main()
