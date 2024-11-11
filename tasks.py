from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import pandas as pd
from RPA.PDF import PDF
from pathlib import Path
import shutil

@task
def minimal_task():
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    close_annoying_modal()
    download_excel_file()
    data = read_excel_file()
    fill_the_form(data)
    archive_receipts()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
def read_excel_file():
    excel = Files()
    df = pd.read_csv("orders.csv")
    # worksheet = excel.read_worksheet_as_table("data", header=True)
    # excel.close_workbook()
    return df

def fill_the_form(dt):
     for index,row in dt.iterrows():
        page = browser.page()
        page.select_option("#head", str(row['Head']))
        page.fill(".form-control", str(row['Legs']))
        page.fill("#address", row['Address'])
        selctr = f"#id-body-{row['Body']}"
        page.click(selctr)
        page.click("#order")
        store_receipt_as_pdf(row['Order number'])
        screenshot_robot(row['Order number'])
        target_path = f"output/final_receipts/{row['Order number']}.pdf"
        pdf_file_path = f"output/receipts/{row['Order number']}.pdf"
        screenshot_path = f"output/screenshots/{row['Order number']}.png"
        embed_screenshot_to_receipt(target_path ,pdf_file_path ,screenshot_path)

def close_annoying_modal():
    page = browser.page()
    page.click("text=OK")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    order_results_html = page.locator("#order-completion").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(order_results_html, f"output/receipts/{order_number}.pdf")

def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path=f"output/screenshots/{order_number}.png")
    page.click("#order-another")
    close_annoying_modal()

def embed_screenshot_to_receipt(target_path, pdf_file,screenshot):
    pdf=PDF()

    list_of_files = [pdf_file,screenshot]
    pdf.add_files_to_pdf(files=list_of_files, target_document=target_path)

def archive_receipts():
    pdf_dir = Path("output/final_receipts")
    output_zip = "output/result"
    shutil.make_archive(output_zip.replace('.zip', ''), 'zip', pdf_dir)
    print(f"Created zip file {output_zip} containing PDF files from {pdf_dir}")
    
     
    