from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

page = browser.page()
http = HTTP()
tables = Tables()
pdf = PDF()
archive = Archive()

@task
def order_robots_from_RobotSpareBin():
    """ Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images. """
    open_robot_order_website()
    orders = get_orders()
    dimensions = tables.get_table_dimensions(orders)
    for row_index in range(dimensions[0]):
        close_annoying_modal()
        order_number = tables.get_table_cell(orders, row_index, 0)
        fill_the_form(row_index, orders)
        pdf_file = store_receipt_as_pdf(order_number)
        screenshot_file = screenshot_robot(order_number)
        embed_screenshot_to_receipt(screenshot_file, pdf_file)
        back_to_the_form()
    archive_receipts()

def open_robot_order_website():
    """ Open the browser and navigate to the given URL """
    browser.configure(browser_engine="chrome", headless=True)
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """ Download the order file and return the table """
    tables = Tables()
    http.download(url="https://robotsparebinindustries.com/orders.csv", target_file="./output/orders.csv", overwrite=True)
    orders = tables.read_table_from_csv(path="./output/orders.csv", header=True)
    return orders

def close_annoying_modal():
    """ Close the modal constitutional rights """
    page.click("button.btn.btn-dark")

def fill_the_form(row_index, orders):
    """ Fill out the form even though the website returns an error """
    tables = Tables()
    head = tables.get_table_cell(table=orders, row=row_index, column="Head")
    body = tables.get_table_cell(table=orders, row=row_index, column="Body")
    legs = tables.get_table_cell(table=orders, row=row_index, column="Legs")
    Address = tables.get_table_cell(orders, row=row_index, column="Address")
    page.select_option("select#head", value=f"{head}")
    page.click("input[value='"f"{body}""']")
    page.click("input[placeholder='Enter the part number for the legs']")
    page.keyboard.insert_text(f"{legs}")
    page.click("#address")
    page.keyboard.insert_text(f"{Address}")
    page.click("#order")
    page.locator("img[alt='Legs']").wait_for()
    result = page.locator("#order-another").is_visible()
    while  result == False:
        page.click("#order")
        page.locator("img[alt='Legs']").wait_for()
        result = page.locator("#order-another").is_visible()
    
def back_to_the_form():
    """ Back to the form to fill again """
    page.click("#order-another")
    page.locator("select#head").wait_for()

def store_receipt_as_pdf(order_number):  
    """ Create a PDF file with receipt number """
    receipt_html = page.locator("#receipt").inner_html()
    pdf.html_to_pdf(receipt_html, f"./output/receipts/{order_number}.pdf") 
    return(f"./output/receipts/{order_number}.pdf")

def screenshot_robot(order_number):
    """ Take a screenshot to the robot """
    page.locator("#robot-preview-image").screenshot(path=f"./output/receipts/{order_number}.png")
    return(f"./output/receipts/{order_number}.png")

def embed_screenshot_to_receipt(screenshot_file, pdf_file):
    """ Embed the robot screenshot to the receipt PDF file """
    list_of_files = [pdf_file, screenshot_file]
    pdf.add_files_to_pdf(list_of_files, pdf_file, False)

def archive_receipts():
    """ Make a ZIP file with PDF archives """
    archive.archive_folder_with_zip(folder="./output/receipts/", archive_name="./output/receipts.zip", include='*.pdf')