from selenium import webdriver # (1) login CDAS
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Chrome()
#driver.get("https://az3.ondemand.esker.com/ondemand/webaccess/asf/home.aspx")
driver.get("https://invoice.eservices.cdas.link/login")
driver.maximize_window()
time.sleep(3)

import pyautogui
pyautogui.moveTo(520, 485, duration=1.5)
pyautogui.click(button='left')
pyautogui.typewrite("")
pyautogui.press('tab')
pyautogui.typewrite(")
pyautogui.press('enter')
time.sleep(1)

try:
    all_invoices = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[1]/aside/div/div/a[2]/div[3]')
    all_invoices.click()
    time.sleep(0.5)
except Exception as e:
    print(e)

def hover(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover.perform()

def hover_click(driver, x_path):
    elem_to_hover = driver.find_element(By.XPATH, x_path)
    #hover = ActionChains(driver).move_to_element(elem_to_hover)
    hover = ActionChains(driver).click(elem_to_hover)
    hover.perform()
time.sleep(1)

import pandas as pd
import openpyxl
import datetime, time, os
from datetime import datetime, timedelta
from tkinter import *
from tkinter import ttk

def remove_pdf_files_within_n_hours(n, path_pdf, end_with): 
    now = datetime.today()
    two_hours_ago = now - timedelta(hours=n)
    #one_day = datetime.timedelta(days=1)
    #before_30_days = today - 30*one_day
    
    for root, _, filenames in os.walk(path_pdf):
        for filename in filenames:
            if filename.lower().endswith('.pdf') and not filename.endswith(end_with):
                file_path = os.path.join(root, filename)
                created_timestamp = os.path.getctime(file_path)
                created_datetime = datetime.fromtimestamp(created_timestamp)
                #created_date = created_datetime.date()
                #created_day = created_datetime.day
                #print(created_datetime)
        
                if created_datetime > two_hours_ago:   # deleting within 5 hours
                    #day_of_week = created_datetime.weekday()    # monday = 0
                    #day_of_week = created_datetime.isoweekday() # monday = 1
                    os.remove(file_path)


global username
username = 'rpa.uat'
def main():
    def divide_by_25(n):
        """
        Divide a number by 25 and return the whole number.
        Args:
            n (int): The number to divide.
        Returns:
            int: The whole number result of the division.
        """
        return n // 25   
   
   
    global username
    username = 'rpa.uat' #get_username_date_to_download()[0]
    
    path_page = "C:/Users/"+username+"/Downloads/cdas_merged/"
    df_page = pd.read_excel(path_page+ r'cdas_page.xlsx', sheet_name='page', engine='openpyxl')
    current_page = df_page['current_page'][0]
    total_page = df_page['total_page'][0]

    import re
    if current_page==0 and total_page ==0:
        numbr_of_pages=driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[3]/div[3]/span')
        match = re.search(r'of (\d+)', numbr_of_pages.get_attribute("textContent"))
        if match:
                total_page = divide_by_25(int(match.group(1)))
                df_page['total_page'] = total_page
                with pd.ExcelWriter(path_page+ r'cdas_page.xlsx', engine='openpyxl') as writer_page:
                    df_page.to_excel(writer_page, sheet_name='page', index=False)
    else:
                current_page = df_page['current_page'][0]
        
    for _ in range(current_page):
        try:        
                right_chevron = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[3]/div[3]/button[3]/span[2]/span/i')
                right_chevron.click()
                time.sleep(1)                
        except Exception as e:
                print(e)
                break
    bill_process_code()
    
    current_page += 1
    df_page['current_page'] = current_page
    with pd.ExcelWriter(path_page+ r'cdas_page.xlsx', engine='openpyxl') as writer_page:
        df_page.to_excel(writer_page, sheet_name='page', index=False)

    remove_pdf_files_within_n_hours(3, r"C:/Users/"+username+r"Downloads/", "_merged.pdf")
   
    windw = Tk() #Create an instance of Tkinter frame
    windw.geometry("550x290") #Set the geometry of Tkinter frame

    def open_popup():
        top= Toplevel(windw)
        top.geometry("550x290")
        top.title("CDAS Automation App")
        Label(top, text= "Thank You!", font=('Mistral 28 bold')).place(x=150,y=80)

    Label(windw, text="CDAS Automated Process Completed.", font=('Helvetica 15 bold')).pack(pady=20)
    #Create a button in the main Window to open the popup
    ttk.Button(windw, text= "Done", command= open_popup).pack()
    windw.mainloop()

from datetime import datetime, timedelta

def bill_process_code():

    def advanced_filter_calendar(driver):

        x_path_hover = '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div' #Advanced filter
        hover(driver, x_path_hover)
        time.sleep(0.5)
        x_path_hover_click = '//*[@id="q-app"]/div/div[2]/div/div[2]/div[2]/div/div'
        hover_click(driver, x_path_hover_click)

        pyautogui.moveTo(360,535, duration=1.5) ##calendar icon
        pyautogui.click(button='left')
        time.sleep(0.5)


    import datetime
    import tkinter as tk
    from tkinter import ttk

    def yesterday(frmt='%Y%m%d', string=True):
        yesterday = datetime.datetime.now() - timedelta(1)
        yesterday_day = yesterday.day
        #yesterday_month = yesterday.month
        return yesterday_day

    
    def get_username_date_to_download():
        window = tk.Tk()
        window.title("CDAS User Input")

        input_labels = ["Username", "Date_To_Download yyyymmdd"]
        for i, label in enumerate(input_labels):
            ttk.Label(window, text=label).grid(row=0, column=i, padx=10, pady=5, sticky='s')

        entries = {'username': [], 'date_to_download': []} #, 'start_time': [], 'finish_time': []
        list_username_date_to_download =[]

        def get_user_input():
            for row in range(len(entries['username'])):
                username = entries['username'][row].get()
                date_to_download = entries['date_to_download'][row].get()
                #start_time = entries['start_time'][row].get()
                #finish_time = entries['finish_time'][row].get()
                #print(f"{username}, {date_to_download}") #{start_time}, {finish_time}")

        for row in range(1, 2):
            username_var = tk.StringVar()
            username_entry = ttk.Entry(window, textvariable=username_var)
            username_entry.grid(row=row, column=0, padx=10, pady=5)
            entries['username'].append(username_var)

            date_to_download_var = tk.StringVar()
            date_to_download_entry = ttk.Entry(window, textvariable=date_to_download_var)
            date_to_download_entry.grid(row=row, column=1, padx=10, pady=5)
            entries['date_to_download'].append(date_to_download_var)

        submit_button = tk.Button(window, text="Submit", command=get_user_input)
        submit_button.grid(row=14, column=0, columnspan=2, pady=10)

        window.mainloop()
        list_username_date_to_download.append(entries['username'][0].get())
        list_username_date_to_download.append(entries['date_to_download'][0].get())
        
        return list_username_date_to_download

    #date_to_download = str(yesterday(frmt='%Y%m%d', string=True))
    date_to_download = str(get_username_date_to_download()[1][-2:])
    #print(date_to_download)
    time.sleep(2)
    
    idx_to_click = str(int(date_to_download) + 6)
    x_path_date = '/html/body/div[3]/div/div[2]/div[1]/div/div[3]/div/div['+idx_to_click+']/button/span[2]/span'

    advanced_filter_calendar(driver)
    time.sleep(0.5)
    date = driver.find_element(By.XPATH, x_path_date)
    if date.get_attribute('textContent') == date_to_download:
        date.click()
        date.click()        
        time.sleep(1.5)

    btn_close=driver.find_element(By.XPATH, '/html/body/div[3]/div/div[2]/div[2]/div[3]/button[2]/span[2]/span')
    btn_close.click()
    time.sleep(0.5)
    pyautogui.press('pagedown')
    time.sleep(0.5)
    #login

    list_invoice_date = [] #get list of invoice date
    i=1
    for _ in range(1, 25):
        #i=1
        try:
            invoice_date_row=driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr['+str(i)+']/td[13]')
            list_invoice_date.append(invoice_date_row.get_attribute("textContent"))   
        except Exception as e:
            print(e)
            break
        i +=1

    ist_bill_ref = []
    list_document_ref = []
    list_bill_ref_saved = []

    import mss
    import mss.tools

    def get_screenshot(path_output):
        with mss.mss() as sct:
            # The screen part to capture
            monitor = {"top": 150, "left": 965, "width": 200, "height": 50}
    
            sct_img = sct.grab(monitor) # Grab the data
            # Save to the picture file
            mss.tools.to_png(sct_img.rgb, sct_img.size, output=path_output)

    path_output = r"C:/Users/"+username+r"/Downloads/cdas_merged/screenshot.png"
    
    from PIL import Image
    import pytesseract
    import re

    #define function to extract text from image
    def extract_text(image_path):
        # Set the path to the Tesseract executable (only required on Windows)
        pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
        image = Image.open(image_path) # Open the image using PIL

        # Preprocess the image (convert to grayscale and apply thresholding)
        image = image.convert('L')  # Convert to grayscale
        image = image.point(lambda x: 0 if x < 128 else 255, '1')  # Apply thresholding

        text = pytesseract.image_to_string(image) # Extract text from the image
        # Use regex to remove special characters like '+' and newline characters
        cleaned_text = re.sub(r'[+\n]', '', text).strip()
        return cleaned_text

    from pathlib import Path
    import pathlib
    
    screenshot=0
    def check_screenshot_exists(path_output):
        if Path(path_output).exists():
            #print("screenshot file exists")
            #screenshot += 1
            time.sleep(0.5)
        else:
            get_screenshot(path_output)
            extract_text(path_output)
            list_save_as_pdf = ["Save as PDF", "Microsoft Print to PDF"]
            if extract_text(path_output) not in list_save_as_pdf:            
                pyautogui.press('down')
    
    def print_save_pdf(bill_ref):
        try:
            pyautogui.moveTo(1050, 165, duration=1.5)
            pyautogui.click(button='left')


            check_screenshot_exists(path_output)
            #pyautogui.press('down')
            pyautogui.press('return')
            time.sleep(0.5)
            for _ in range(3):
                pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('return')
            time.sleep(0.5)

            pyautogui.press('delete')
            time.sleep(0.5)
            pyautogui.typewrite(bill_ref+ '_bill')
            pyautogui.press('return')
            time.sleep(0.5)
            list_bill_ref_saved.append(bill_ref)
        except Exception as e:
            print(f'pdf not saved {e}')
            #break
        return list_bill_ref_saved
    
    from PyPDF2 import PdfReader, PdfWriter
    import fitz
    import os
    #import pathlib

    def get_pdf_files_with_invoice_number(path_pdf, bill_ref):
        """
        Gets a list of PDF files in the given directory that contain the specified invoice number in their filename.
        Args:
            path_pdf: The path to the directory containing the PDF files.
            invoice_numbr: The invoice number to search for.
        Returns:
            A list of PDF filenames that contain the invoice number.
        """
        pdf_files = [f for f in os.listdir(path_pdf) if bill_ref in f and f.endswith(".pdf")]
        return pdf_files
    
    
    def merge_pdfs(pdf_files, path_pdf):
        """
        Merges multiple PDF files into a single PDF.
        Args:
            pdf_files: A list of PDF filenames.
            path_pdf: The path to the directory containing the PDF files.
        Returns:
            The merged PDF writer object.
        """
        pdf_writer = PdfWriter()

        for pdf in pdf_files:
            try:
                with open(path_pdf + pdf, 'rb') as pdf_file:
                    pdf_reader = PdfReader(pdf_file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        pdf_writer.add_page(page)
            except EOFError as e:
                #print(f"Error reading PDF file '{pdf}': {e}")
                continue  # Skip the file that caused the error
        return pdf_writer
    
    import glob, os.path
    #import pathlib
    #from pathlib import Path
    import shutil
    
    # define function move pdf
    def move_pdf(src_dir, dest_dir):
        """
        Moves a PDF file from the source directory to the destination directory.
        Args:
            src_dir (str): The full path to the source directory containing the PDF file.
            dest_dir (str): The full path to the destination directory where the PDF file will be moved.
        Returns:
            str: A message indicating success or failure.
        """
        try:
            # Check if the source directory exists
            if not os.path.exists(src_dir):
                    return f"Error: Source directory '{src_dir}' does not exist."

            # Check if the destination directory exists; if not, create it
            if not os.path.exists(dest_dir):
                    os.makedirs(dest_dir)  # Create the destination directory if it doesn't exist
                    print(f"Info: Destination directory '{dest_dir}' created.")

            # Find the first PDF file in the source directory
            pdf_files = [f for f in os.listdir(src_dir) if (f.lower().endswith('.pdf') and 'merged' in f.lower())]
            if not pdf_files:
                    return f"Error: No PDF files found in the source directory '{src_dir}'."

            # Get the first PDF file (assuming there's only one PDF file to move)
            for pdf_file in pdf_files:
                    pdf_file_name = pdf_file#pdf_files[0]
                    src_file_path = os.path.join(src_dir, pdf_file_name)
                    dest_file_path = os.path.join(dest_dir, pdf_file_name)

                    # Check if a file with the same name already exists in the destination directory
                    if os.path.exists(dest_file_path):
                        return f"Error: A file with the name '{pdf_file_name}' already exists in '{dest_dir}'."

                    # Copy and Move the file
                    shutil.copy(src_file_path, dest_file_path)
                    shutil.move(src_file_path, dest_file_path)
                    return f"Success: File '{pdf_file_name}' moved from '{src_dir}' to '{dest_dir}'."

        except PermissionError:
                return f"Error: Permission denied. Check if you have the necessary permissions to access '{src_dir}' or '{dest_dir}'."
        except FileNotFoundError:
                return f"Error: The file '{pdf_file_name}' was not found in '{src_dir}'."
        except Exception as e:
                return f"Error: An unexpected error occurred - {str(e)}"

    #username = 'rpa.uat'
    path_pdf = r"C:/Users/"+ username+ r"/Downloads/" ##len(list_invoice_date)
    fldr = path_pdf+ r"cdas_merged/"
    if os.path.isdir(fldr):
                    #print(f'{fldr} folder exists')
                    time.sleep(0.5)
    else:
                    os.mkdir(fldr)
    
    list_bill_ref = []
    for i in range(1, 3):
        advanced_filter_calendar(driver)
        time.sleep(0.5)
        pyautogui.press('pagedown')
        time.sleep(0.5)
        try:
            bill_ref_row = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr['+str(i)+']/td[4]')
            bill_ref = bill_ref_row.get_attribute("innerHTML")
            list_bill_ref.append(bill_ref)
            
        except Exception as e:
            print(e)
            break

        try:
            documents = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr['+str(i)+']/td[5]/div/div/div/div/a')
            list_document_ref.append(documents.get_attribute("innerHTML"))
            documents.click() #download Documents attachments
            time.sleep(1)
            pyautogui.moveTo(480,170, duration=1.5)
            pyautogui.click(button='left')
            time.sleep(1)
        except Exception as e:
            #print(e)
            time.sleep(0.5)
        try:
        
                view = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div[2]/table/tbody/tr['+str(i)+']/td[1]/button/span[2]')
                view.click()
                time.sleep(0.5)
        except Exception as e:
                print(e)
                break
        time.sleep(1)

        try:
                print = driver.find_element(By.XPATH, '//*[@id="q-app"]/div/div[2]/div/div[2]/div[3]/div/div/button/span[2]')
                print.click()
                time.sleep(0.5)
        except Exception as e:
                print(e)
                break

        try:
                print_save_pdf(bill_ref)
        except Exception as e:
                print(e)
                break

        driver.back() #page back
        time.sleep(2)

        try:        
                #bill_ref = 'GI25026078'
                get_pdf_files_with_invoice_number(path_pdf, bill_ref)

                merged_writer_bill=merge_pdfs(get_pdf_files_with_invoice_number(path_pdf, bill_ref), path_pdf)
        
                with open(path_pdf+bill_ref+"_merged.pdf", "wb") as output_file:
                    merged_writer_bill.write(output_file) # Save the merged PDF 
                    #print(f'merge pdf {bill_ref} saved')
        except Exception as e:
                #print(f'merge pdf not saved {e}')
                time.sleep(0.5)
        time.sleep(1)
    
        try:
                move_pdf(path_pdf, fldr)
                time.sleep(1)
        except Exception as e:
                print(e)
                break

        
        
    #run

if __name__ == "__main__":
    main()

