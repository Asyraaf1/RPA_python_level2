from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
# from selenium import webdriver
# from selenium.webdriver.common.by import By
import time
# from reportlab.pdfgen import canvas
# import pyautogui
from tkinter import *
import pyscreenshot
from tkinter.filedialog import *
import pdfkit
from PyPDF2 import PdfMerger
from PIL import Image


path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates a ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )

    open_robot_order_website()
    download_excel_file()
    fill_form_with_excel_data()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page = browser.page()
    page.click("button:text('ok')")

def download_excel_file():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_form_with_excel_data():
    excel = Tables()
    worksheet = excel.read_table_from_csv("orders.csv")
    for row in worksheet:
        fill_and_submit_sales_form(row)

def fill_and_submit_sales_form(sales_rep):
    page = browser.page()
    
    page.select_option("#head", str(sales_rep['Head']))
    body_radio_button_selector = f'input[name="body"][value="{str(sales_rep["Body"])}"]'
    page.check(body_radio_button_selector)
    page.fill(".form-control", str(sales_rep['Legs']))
    page.fill("#address", sales_rep['Address'])
    page.click("button:text('order')")
    time.sleep(2)
    # is_element_visible()

    # if  webdriver.Chrome(page).find_element(By.CLASS_NAME, "alert-danger") == True:
    #     print("Element is visible")
    #     page.click("button:text('order')")
    #     page.click("button:text('Order another robot')")
    #     page.click("button:text('ok')")
    # else:
    #     print("Element is not visible")
    #     page.click("button:text('Order another robot')")
    #     page.click("button:text('ok')")

    # if is_element_visible(page, ".alert-danger"):
    #     print("Element is visible")
    #     fill_and_submit_sales_form(sales_rep)

    if is_element_visible(page, "//div[@class='alert alert-danger']"):
        print("Element is visible")
        fill_and_submit_sales_form(sales_rep)

    else:
        print("Element is not visible")
        screenshot_robot(sales_rep["Order number"])
        save_page_as_pdf(sales_rep["Order number"])
        merge_image_with_pdf("output/pdf_folder/"+sales_rep["Order number"]+"_receipt.pdf","output/screenshot_folder/"+sales_rep["Order number"]+"_screenshot.png", "output/merged_folder/"+sales_rep["Order number"]+"_merged.pdf")
        page.click("button:text('Order another robot')")
        page.click("button:text('ok')")


# def is_element_visible(page, selector):
#     try:
#         return page.query_selector(selector) is not None
#     except Exception:
#         return False
        
def is_element_visible(page, xpath):
    # Check if an element matching the provided selector is visible on the page
    try:
        return page.query_selector(xpath) is not None
    except Exception:
        return False
    
def screenshot_robot(order_number):
    ss=pyscreenshot.grab()
    # ss_save=asksaveasfilename()
    # ss.save(ss_save+"_screenshot.png")

    ss_save = ("output/screenshot_folder/"+order_number+"_screenshot.png" ) # Provide the desired filename and path
    ss.save(ss_save)
    
def save_page_as_pdf(order_number):
    page_source = browser.page().inner_html("#receipt")
    output_pdf_path = f"output/pdf_folder/{order_number}_receipt.pdf"  # Specify the output file path
    pdfkit.from_string(page_source, output_pdf_path, configuration=pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf))



#      # Capture screenshot
#     screenshot = pyautogui.screenshot()
    
#     # Save screenshot
#     screenshot.save('receipt_screenshot.png')
    
#     # Convert receipt to PDF
#     with open('receipt.txt', 'r') as f:
#         receipt_text = f.read()
    
#     # Create PDF file
#     pdf_filename = 'receipt.pdf'
#     c = canvas.Canvas(pdf_filename)
#     c.drawString(100, 750, receipt_text)
#     c.save()
    
#     print("Receipt PDF and screenshot created successfully.")
    
def merge_image_with_pdf(input_pdf_path, image_path, output_pdf_path):
    merger = PdfMerger()

    # Open the input PDF
    with open(input_pdf_path, 'rb') as input_pdf:
        merger.append(input_pdf)

    # Open the image file
    image = Image.open(image_path)

    # Convert the image to PDF format
    image_pdf_path = 'image.pdf'
    image.save(image_pdf_path, 'PDF')

    # Merge the image PDF with the input PDF
    with open(image_pdf_path, 'rb') as image_pdf:
        merger.append(image_pdf)

    # Write the merged PDF to the output file
    with open(output_pdf_path, 'wb') as output_pdf:
        merger.write(output_pdf)

    # # Clean up the temporary image PDF file
    # if os.path.exists(image_pdf_path):
    #     os.remove(image_pdf_path)
    
# only uncomment if run using debugger(f5)
order_robots_from_RobotSpareBin()