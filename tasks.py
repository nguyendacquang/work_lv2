import csv
import time
import os
from zipfile import ZipFile
from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.PDF import PDF

@task
def open_robot_order_website():
    download_excel_file()
    orders()
    time.sleep(10)
    
def orders():
    browser = Selenium()
    browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order")
    browser.maximize_browser_window()
    os.makedirs("output/img")
    os.makedirs("output/pdf")
    with open("output/orders.csv", mode='r', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        header = next(csv_reader)  # Skip the header row if there is one
        for row in csv_reader:
            print(row)
            browser.wait_until_element_is_visible("//button[normalize-space()='OK']", timeout=10)
            browser.click_element("//button[normalize-space()='OK']")
            browser.select_from_list_by_value("//select[@id='head']", row[1])
            browser.click_element(f"//input[@value={row[2]}]")
            browser.input_text("//*[@placeholder='Enter the part number for the legs']", row[3])
            browser.input_text("//*[@id='address']", row[4])
            browser.wait_until_element_is_visible("//button[@id='order']", timeout=10)
            browser.click_element("//button[@id='preview']")
            time.sleep(1)   
            browser.screenshot(locator="//*[@id='robot-preview-image']", filename=f"output/img/{row[0]}.png")
            end_time = time.time() + 30
            while time.time() < end_time:
                browser.execute_javascript("document.querySelector('button[id=\"order\"]').click();")
                try:
                    # Try to find the new element with a short timeout to avoid long waits
                    browser.wait_until_page_contains_element("//button[@id='order-another']", timeout=2)
                    print("New element appeared!")
                    break
                except Exception as e:
                    # If the new element is not found, continue retrying
                    print("New element not found, retrying...")
                    time.sleep(1)  # Wait a bit before retrying to avoid too frequent attempts
            browser.wait_until_element_is_visible("//*[@id='receipt']", timeout=10)
            sales_results_html = browser.get_element_attribute("//*[@id='receipt']", "innerHTML")
            create_pdf_from_html(sales_results_html, f"output/pdf/{row[0]}.pdf")
            embed_screenshot_in_pdf(f"output/img/{row[0]}.png", f"output/pdf/{row[0]}.pdf", f"output/pdf/merge_{row[0]}.pdf")
            browser.click_element("//button[@id='order-another']")

    browser.close_browser()
    zip_folder("output/pdf", "output/pdf.zip")

def embed_screenshot_in_pdf(image_path, source_pdf, output_pdf):
    pdf = PDF()
    # Thêm ảnh chụp màn hình vào tệp PDF
    pdf.add_watermark_image_to_pdf(
        image_path=image_path,
        source_path=source_pdf,
        output_path=output_pdf,
        coverage=0.2  # Điều chỉnh kích thước ảnh trên trang PDF, giá trị mặc định là 0.2
    )

def create_pdf_from_html(html_content_as_string, output_file_path):
    pdf = PDF()
    # Convert HTML to PDF
    pdf.html_to_pdf(html_content_as_string, output_file_path)

def download_excel_file():
    """Downloads CSV file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",target_file="output", overwrite=True)

def zip_folder(folder_path, output_path):
    with ZipFile(output_path, 'w') as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                zipf.write(os.path.join(root, file), 
                           os.path.relpath(os.path.join(root, file), 
                                           os.path.join(folder_path, '..')))