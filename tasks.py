from datetime import datetime

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=100,
    )
    open_robot_order_website()
    orders = get_orders()
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
    archive_receipts()
    
def open_robot_order_website():
    """Opens the order website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")

def get_orders():
    """Gets the orders from the .csv"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)

    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv",columns=["Order number", "Head","Body","Legs","Address"]
    )
    return orders

def fill_the_form(row):
    page = browser.page()
    page.select_option("#head",row['Head'])
    page.check("#id-body-"+row['Body'])
    page.fill("xpath=//label[contains(.,'3. Legs:')]/../input",row['Legs'])
    page.fill("#address",row['Address'])
    page.click("#preview")
    page.click("#order")
    failed = page.is_visible("#order")
    i = 0
    while(failed and i < 3):
        page.click("#order")
        failed = page.is_visible("#order")
        i = i + 1

    pdf_file = store_receipt_as_pdf(row['Order number'])
    screenshot = screenshot_robot(row['Order number'])
    embed_screenshot_to_receipt(screenshot, pdf_file)

    page.click("#order-another")

def store_receipt_as_pdf(order_number):
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdfpath = "output/receipts/"+order_number+".pdf"
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, pdfpath)
    return pdfpath

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    filepath = "output/screenshots/"+order_number+".png"
    page.screenshot(path=filepath)
    return filepath

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    files_to_append = [screenshot]
    pdf.add_files_to_pdf(files=files_to_append,target_document=pdf_file,append=True)

def archive_receipts():
    current_time = datetime.now().strftime("%d-%m-%Y-%H-%M")
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts/",f"receipts_{current_time}")
    filesystem = FileSystem()
    filesystem.remove_directory("output/screenshots/",True)
    filesystem.remove_directory("output/receipts/",True)