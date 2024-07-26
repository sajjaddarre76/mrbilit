from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import time
import base64
import requests
from dotenv import load_dotenv
import os
import convert_numbers

# captcha code recognition
from PIL import Image, ImageEnhance
import cv2
import pytesseract
import numpy as np
import pandas as pd

########################################## FUNCTIONS ##########################################################

# getting the data from env file (origin, destination, departure_date, destination_date, phonenumber, password, driver_path)
def get_primary_information():
    # Load environment variables from the .env file
    load_dotenv(".env", override=True)

    # Access the variables
    user_phonenumber = os.getenv('PHONENUMBER')
    user_password = os.getenv('PASSWORD')
    driver_path = os.getenv('DRIVERPATH')
    origin = os.getenv('ORIGIN')
    destination = os.getenv('DESTINATION')
    departure_date = os.getenv('DEPARTUREDATE')
    return_date = os.getenv('RETURNDATE')
    passengers_details_excel_path = os.getenv('PASSENGERSEXCELFILE')
    bed_count = os.getenv('BEDCOUNT')
    cardnumber = os.getenv('CARDNUMBER')
    cvv2 = os.getenv('CVV2')
    card_expiration_year = os.getenv('CARDEXPIRATIONYEAR')
    card_expiration_month = os.getenv('CARDEXPIRATIONMONTH')
    captcha_recognition = os.getenv('CAPTCHARECOGNITION')
    return user_phonenumber, user_password, driver_path, origin, destination, departure_date, return_date, passengers_details_excel_path, bed_count, cardnumber, cvv2, card_expiration_year, card_expiration_month, captcha_recognition

# date manipulation
def shamsi_month(month_number):
    shamsi_months = {"1": "فروردین", "2": "اردیبهشت", "3": "خرداد", "4": "تیر", "5": "مرداد", "6": "شهریور", "7": "مهر", "8": "آبان", "9": "آذر", "10": "دی", "11": "بهمن", "12": "اسفند"}
    return shamsi_months[convert_numbers.persian_to_english(month_number)]

# Excel file manipulation
def get_passenger_excel(file_path, bed_count):
    df = pd.read_excel(file_path)
    # Convert all values in the DataFrame to strings
    df = df.apply(lambda x: x.astype(str))
    request_turn = int(len(df) / int(bed_count))
    if len(df) % int(bed_count) != 0:
        request_turn += 1
    if request_turn == 0:
        request_turn = 1
    return request_turn, df

def iterate_over_chunk_of_passengers(chunk):
    
    return chunk.iloc[:, 0].notnull().all()

def process_excel_file(df, chunk_size=4):
    """
    Reads an Excel file, processes it in chunks of specified row size,
    and removes successfully processed rows from the DataFrame.
    
    Args:
        file_path (str): Path to the input Excel file.
        chunk_size (int): Number of rows per chunk. Default is 4.
    """
    # Read the Excel file
    
    while not df.empty:
        # Get the first chunk of rows
        chunk = df.iloc[:chunk_size]

        # Process the chunk
        if process_rows(chunk):
            # If the process is successful, remove the chunk from the DataFrame
            df = df.iloc[chunk_size:]
            print(f"Processed and removed chunk:\n{chunk}")
        else:
            # If the process fails, handle the failure (e.g., log error, stop processing)
            print(f"Failed to process chunk:\n{chunk}")
            break

    # Save the remaining DataFrame back to the same Excel file
    df.to_excel(file_path, index=False)
    print(f"Remaining data saved to '{file_path}'")

# TimeoutException handling
def find_element_with_retry(driver, locator, retries=3, delay=5, url=None, multiple=False, driverWait=10, button=None, inputField=None):
    attempt = 0
    while attempt < retries:
        try:
            if attempt > 0 and url:
                # Re-request the URL before retrying
                print(f"Retrying URL request: {url}")
                driver.get(url)
            if multiple:
                element = WebDriverWait(driver, driverWait, poll_frequency=0.5).until(
                    EC.visibility_of_all_elements_located(locator)
                )
            elif button:
                element = WebDriverWait(driver, driverWait, poll_frequency=0.5).until(
                    EC.element_to_be_clickable(locator)
                )
            elif inputField:
                element = WebDriverWait(driver, driverWait, poll_frequency=0.5).until(
                    EC.visibility_of_element_located(locator)
                )
            else:
                element = WebDriverWait(driver, driverWait, poll_frequency=0.5).until(
                    EC.presence_of_element_located(locator)
                )
                
            return element
        except TimeoutException:
            print(f"Attempt {attempt + 1} of {retries} failed: TimeoutException")
            attempt += 1
            time.sleep(delay)  # Wait before the next retry
    
    # If the element is not found after all retries, raise an exception
    raise TimeoutException(f"Failed to find element after {retries} retries")

# checking for images to be fully loaded
# Define a function to check if the image is completely loaded
def is_image_loaded(driver, image_locator):
    try:
        # Wait for the image element to be present
        image = WebDriverWait(driver, 10, poll_frequency=0.5).until(EC.presence_of_element_located(image_locator))
        
        # Use JavaScript to check if the image is completely loaded
        is_loaded = driver.execute_script(
            "return arguments[0].complete && "
            "typeof arguments[0].naturalWidth != 'undefined' && arguments[0].naturalWidth > 0",
            image
        )
        return is_loaded

    except Exception as e:
        print(f"An error occurred: {e}")
        return False

# Define a function to check if the condition is met
def check_condition(locator, driverWait=10):
    try:
        # Example condition: Check if a specific element is present after submitting the code
        WebDriverWait(driver, driverWait, poll_frequency=0.5).until(EC.presence_of_element_located(locator))  # Replace with the actual locator
        return True
    except:
        return False

# Function to click a button and wait for it to disappear
def wait_for_button_disappearance(button_locator, timeout=10):
    # Wait for the button to disappear
    WebDriverWait(driver, timeout).until(EC.invisibility_of_element_located(button_locator))

def Initialize_driver():
    # Path to your Chrome user data directory
    # user_data_dir = r"C:\Users\edr\Desktop\Projects and ideas\Ticket_Reservation_automation_project\mrbilit_bot\selenium-profile"
    chrome_options = Options()
    # chrome_options.add_argument(f"user-data-dir={user_data_dir}")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    # driver = webdriver.Chrome(service=ChromeService(driver_path), options=chrome_options)
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver

# Image enhancement for captcha recognition
# Set up the path to tesseract executable
def preprocess_image(image_path):
    # Open the image using PIL
    image = Image.open(image_path)
    
    # Convert the image to grayscale
    image = image.convert('L')
    
    # Enhance the contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    
    # Convert the image to a numpy array
    image_np = np.array(image)
    
    # Apply Gaussian blur to remove noise
    image_np = cv2.GaussianBlur(image_np, (5, 5), 0)
    
    # Apply adaptive thresholding to make the text stand out
    image_np = cv2.adaptiveThreshold(image_np, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    
    # Use morphology operations to remove small noise
    kernel = np.ones((2, 2), np.uint8)
    image_np = cv2.morphologyEx(image_np, cv2.MORPH_OPEN, kernel)
    
    # Save the processed image (optional)
    processed_image_path = 'processed_captcha.png'
    cv2.imwrite(processed_image_path, image_np)
    
    return processed_image_path

def wait_for_image_to_load(driver, image_locator, timeout=20):
    try:
        # Define a wait time
        wait = WebDriverWait(driver, timeout, poll_frequency=0.5)

        # Wait until the image is fully loaded
        wait.until(lambda d: driver.execute_script(
            "return arguments[0].complete && typeof arguments[0].naturalWidth != 'undefined' && arguments[0].naturalWidth > 0",
            d.find_element(*image_locator)
        ))

        return True
    except TimeoutException:
        return False


####################################### FUNCTION CALLS and SCRIPTS ####################################################

# getting the primary information
user_phonenumber, user_password, driver_path, origin, destination, departure_date, return_date, passengers_details_excel_path, bed_count, cardnumber, cvv2, card_expiration_year, card_expiration_month, captcha_recognition = get_primary_information()

# getting the request turn and the dataframe of the passengers details
request_turn, df = get_passenger_excel(passengers_details_excel_path, int(bed_count))

TURN = 1

while TURN <= request_turn:
    turn = TURN * int(bed_count)
    print(TURN, turn, request_turn)
    passengers_details_chunk_df = df.iloc[turn - int(bed_count):turn]
    print(passengers_details_chunk_df)
    passengers_count = len(passengers_details_chunk_df)
    main_url = f"https://mrbilit.com/trains/{origin}-{destination}?departureDate={departure_date}&adultCount={len(passengers_details_chunk_df)}&returnDate={return_date}"
    driver = Initialize_driver()

    try:
        driver.get(main_url)
        # Open a new tab using JavaScript
        driver.execute_script("window.open('');")

        # Switch to the new tab
        driver.switch_to.window(driver.window_handles[1])

        # Open the second URL in the new tab
        second_url = "https://messages.google.com/web/conversations"
        driver.get(second_url)

        # Optionally, interact with the pages
        # Switch back to the first tab
        driver.switch_to.window(driver.window_handles[0])

        # # Switch to the second tab
        # driver.switch_to.window(driver.window_handles[1])
        # print("Second Tab Title:", driver.title)

        # Define the locator for the element we want to find
        tickets = (By.CSS_SELECTOR, ".card-section-wrapper.cards-container")

        # Reserving departure tickets
        reservable_tickets = find_element_with_retry(driver, tickets, retries=5, delay=2, url=main_url, multiple=False)
        reservable_tickets = reservable_tickets.find_elements(By.CSS_SELECTOR, "div.trip-card-container")
        for ticket in reservable_tickets:
            bed_count_element = ticket.find_element(By.CSS_SELECTOR, "span.title")
            capacity_count = ticket.find_element(By.CSS_SELECTOR, "div.capacity-text").text.split(' ')[0]
            train_bed_count = bed_count_element.text
            if ((f"{bed_count} تخته" in train_bed_count) or (f"{bed_count} تخته" in train_bed_count) or (f"{bed_count}تخته" in train_bed_count) or (f"{bed_count}تخته" in train_bed_count)) and (int(capacity_count) >= passengers_count):
                driver.execute_script("arguments[0].scrollIntoView(true);", ticket)
                reserve_button = ticket.find_element(By.CSS_SELECTOR, "button.reserve-button")
                reserve_button.click()
                break
        # Reserving return tickets
        if return_date != None:
            reservable_tickets = find_element_with_retry(driver, tickets, retries=5, delay=2, url=None, multiple=False)
            reservable_tickets = reservable_tickets.find_elements(By.CSS_SELECTOR, "div.trip-card-container")
            for ticket in reservable_tickets:
                bed_count_element = ticket.find_element(By.CSS_SELECTOR, "span.title")
                capacity_count = ticket.find_element(By.CSS_SELECTOR, "div.capacity-text").text.split(' ')[0]
                train_bed_count = bed_count_element.text
                if ((f"{bed_count} تخته" in train_bed_count) or (f"{bed_count} تخته" in train_bed_count) or (f"{bed_count}تخته" in train_bed_count) or (f"{bed_count}تخته" in train_bed_count)) and (int(capacity_count) >= passengers_count):
                    driver.execute_script("arguments[0].scrollIntoView(true);", ticket)
                    reserve_button = ticket.find_element(By.CSS_SELECTOR, "button.reserve-button")
                    reserve_button.click()
                    break
        
        time.sleep(3)
        # دکمه ادامه و درج مشخصات
        button_text = "ادامه و درج مشخصات"
        locator = (By.CSS_SELECTOR, "button.proceed-button")
        proceed_button = find_element_with_retry(driver, locator, retries=5, delay=2, url=None, multiple=False,driverWait=5)
        while not button_text == proceed_button.text:
                proceed_button = find_element_with_retry(driver, locator, retries=5, delay=2, url=None, multiple=False,driverWait=5)
        proceed_button.click()
            



        # وارد کردن اطلاعات مسافران 
        for i in range(0, passengers_count):
            passenger = passengers_details_chunk_df.iloc[i]
            locator = (By.ID, f"user-{i}")
            passenger_details_element = find_element_with_retry(driver, locator, retries=5, delay=20, url=None, multiple=False)
            driver.execute_script("arguments[0].scrollIntoView(true);", passenger_details_element)
            passanger_gender_element = passenger_details_element.find_elements(By.CSS_SELECTOR, "div.checkbox-label")
            if passenger['جنس'] == "مرد":
                passanger_gender_element[0].click()
            else:
                passanger_gender_element[1].click()
            passanger_name_element = passenger_details_element.find_element(By.CSS_SELECTOR, "input[placeholder='نام']")  # Adjust the timeout as necessary
            passanger_name_element.send_keys(passenger['نام'])
            passanger_lastname_element = passenger_details_element.find_element(By.CSS_SELECTOR, "input[placeholder='نام خانوادگی']")  # Adjust the timeout as necessary
            passanger_lastname_element.send_keys(passenger['نام خانوادگی'])
            passenger_passcode_element = passenger_details_element.find_element(By.CSS_SELECTOR, "input[placeholder='شماره ملی']")  # Adjust the timeout as necessary
            passenger_passcode_element.send_keys(passenger['کد ملی'])
            birth_date_segments = passenger_details_element.find_elements(By.CSS_SELECTOR, "div.mr-select-input-container.select-input")
            for i in range(0, len(birth_date_segments)):
                birth_date_segments[i].click()
                choosable_elements = birth_date_segments[i].find_elements(By.CSS_SELECTOR, "div.select-label")
                for item in choosable_elements:
                    if i == 0 and item.text == passenger["روز"]:
                        item.click()
                        break
                    if i == 1 and item.text == shamsi_month(passenger["ماه"]):
                        item.click()
                        break
                    if i == 2 and item.text == passenger["سال"]:
                        item.click()
                        break
        # وارد کردن شماره تلفن
        phonenumber_element = find_element_with_retry(driver, (By.CSS_SELECTOR, "input[placeholder='تلفن همراه']"), retries=5, delay=20, url=None, multiple=False)
        # phonenumber_element = driver.find_element(By.CSS_SELECTOR, "input[placeholder='تلفن همراه']")
        phonenumber_element.send_keys(user_phonenumber)

        # تایید و ادامه
        proceed_button = driver.find_element(By.CSS_SELECTOR, "button.proceed-button")
        driver.execute_script("arguments[0].scrollIntoView(true);", proceed_button)
        proceed_button.click()


        # دکمه تایید و پرداخت
        button_text = "تأیید و پرداخت"
        locator = (By.CSS_SELECTOR, "button.proceed-button")
        proceed_button = find_element_with_retry(driver, locator, retries=5, delay=10, url=None, multiple=False,driverWait=5)
        while not button_text == proceed_button.text:
            proceed_button = find_element_with_retry(driver, locator, retries=5, delay=2, url=None, multiple=False,driverWait=5)
        driver.execute_script("arguments[0].scrollIntoView(true);", proceed_button)
        proceed_button.click()

        # ورود اطلاعات کارت اعتباری
        card_number_element = find_element_with_retry(driver, (By.ID, "CardNumber_PanString"), retries=5, delay=2, url=None, multiple=False)
        card_number_element.send_keys(cardnumber)
        cvv2_element = driver.find_element(By.ID, "Cvv2")
        cvv2_element.send_keys(cvv2)
        card_expiration_month_element = driver.find_element(By.ID, "Month")
        card_expiration_month_element.send_keys(card_expiration_month)
        card_expiration_year_element = driver.find_element(By.ID, "Year")
        card_expiration_year_element.send_keys(card_expiration_year)

        if captcha_recognition == "ON":
            # Captcha code recognition using tesseract
            # Set the path for the Tesseract executable
            pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe' 

            # Locate the CAPTCHA image element
            WebDriverWait(driver, 5)
            capctha_locator = (By.CSS_SELECTOR, "img#CaptchaImage")
            captcha_element = driver.find_element(By.CSS_SELECTOR, "img#CaptchaImage") 
            # Locate the CAPTCHA input field
            captcha_input_element = driver.find_element(By.ID, 'CaptchaInputText')

            # Wait until the image is fully loaded
            if wait_for_image_to_load(driver, capctha_locator):
                # Save the CAPTCHA image
                captcha_image_path = "captcha.png"
                captcha_element.screenshot(captcha_image_path)

                # Preprocess the CAPTCHA image
                processed_image_path = preprocess_image(captcha_image_path)

                # Open the processed image and recognize the text
                captcha_image = Image.open(processed_image_path)
                captcha_text = pytesseract.image_to_string(captcha_image, config='--psm 7')

                # Print the recognized CAPTCHA text
                print(f"Recognized CAPTCHA text: {captcha_text}")

                # Enter the interpreted CAPTCHA text
                captcha_input_element.send_keys(captcha_text.strip())

                captcha_turn = 1
                while captcha_turn < 6:
                    captcha_input_value = captcha_input_element.get_attribute('value').strip()
                    if not captcha_input_value or len(captcha_input_value) != 5:
                        # Preprocess the CAPTCHA image
                        captcha_image = Image.open(processed_image_path)
                        captcha_text = pytesseract.image_to_string(captcha_image)
                        captcha_input_element.send_keys(captcha_text.strip())
                    else:
                        break
                    captcha_turn += 1

            
            # دکمه دریافت رمز پویا
            otp_button = driver.find_element(By.CSS_SELECTOR, "button#Otp")
            otp_button.click()

            # کد امنیتی به درستی وارد نشده است
            security_code_error = find_element_with_retry(driver, (By.XPATH, "//p[text()='کد امنیتی به درستی وارد نشده است']"), driverWait=2)
            if security_code_error:
                capctha_locator = (By.CSS_SELECTOR, "img#CaptchaImage")
                captcha_element = driver.find_element(By.CSS_SELECTOR, "img#CaptchaImage") 
                # Locate the CAPTCHA input field
                captcha_input_element = driver.find_element(By.ID, 'CaptchaInputText')

                # Wait until the image is fully loaded
                if wait_for_image_to_load(driver, capctha_locator):
                    # Save the CAPTCHA image
                    captcha_image_path = "captcha.png"
                    captcha_element.screenshot(captcha_image_path)

                    # Preprocess the CAPTCHA image
                    processed_image_path = preprocess_image(captcha_image_path)

                    # Open the processed image and recognize the text
                    captcha_image = Image.open(processed_image_path)
                    captcha_text = pytesseract.image_to_string(captcha_image, config='--psm 7')

                    # Print the recognized CAPTCHA text
                    print(f"Recognized CAPTCHA text: {captcha_text}")

                    # Enter the interpreted CAPTCHA text
                    captcha_input_element.send_keys(captcha_text.strip())

        # انتظار تا دانلود فایل بلیط
        locator = (By.XPATH, "//button[text()='دانلود فایل بلیط']")
        check_counter = 0
        while not check_condition(locator) and check_counter < 6:
            time.sleep(10)
            check_counter += 1

        if check_condition(locator):
            TURN += 1
            # Define the path to the Excel file
            file_path = 'reserve_status.xlsx'
            passengers_details_chunk_df["رفت"] = pd.NA
            passengers_details_chunk_df["برگشت"] = pd.NA
            # Iterate through each row and set the value for the new column
            for index, row in passengers_details_chunk_df.iterrows():
                # Replace this logic with whatever you need to compute the new column value
                passengers_details_chunk_df.at[index, 'رفت'] = "Y"
                if return_date:
                    passengers_details_chunk_df.at[index, 'برگشت'] = "Y"
                else:
                    passengers_details_chunk_df.at[index, 'برگشت'] = "N"

            # Check if the file exists
            if os.path.exists(file_path):
                df_reserver_status = pd.read_excel(file_path, engine='openpyxl')
                # Concatenate DataFrames to add new rows
                df_reserver_status = pd.concat([df_reserver_status, passengers_details_chunk_df], ignore_index=True)
                # Save the modified DataFrame back to the same Excel file
                df_reserver_status.to_excel(file_path, index=False, engine='openpyxl')
            else: 
                df_reserver_status = pd.DataFrame(passengers_details_chunk_df)
                # Write the DataFrame to a new Excel file
                df_reserver_status.to_excel(file_path, index=False, engine='openpyxl')

        
        download_ticket_btn = driver.find_element(By.XPATH, "//button[text()='دانلود فایل بلیط']")
        download_ticket_btn.click()
        time.sleep(10)
        continue
    except TimeoutException as e:
        print("Error:", e)

    finally:
        driver.quit()