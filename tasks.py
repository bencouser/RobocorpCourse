from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    # Open Website
    browser.configure(slowmo=1000,)
    open_robot_order_website()

    # Navigate to order page (beginning of loop)
    browser.configure(slowmo=100,)
    close_annoying_modal()

    # Get Input Data
    download_excel_file(URL="https://robotsparebinindustries.com/orders.csv")

    # Process Data
    fill_form_with_excel_data()

    # Archive Report
    archive_receipts()

def open_robot_order_website():
    """
    Open browser on robot order website.
    """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def close_annoying_modal():
    """
    Agree to give data away and navigate to beginning of process loop
    """
    page = browser.page()
    page.click("text=OK")

def download_excel_file(URL: str):
    """
    Input: Str URL pointing to data source
    Download input file using HTTP request
    """
    http = HTTP()
    http.download(url=URL, overwrite=True)

def fill_and_submit_form(order):
    """
    Input: Order. A row item from the report that should contain Order number, Head, Body, Legs and Address
    Take input data, fill in form and submit.
    """

    # Map robot names to indexes provided from input data
    # Note that we will need to -1 the indexes to prevent subscript being out of range
    ROBOT_NAMES = ["Roll-a-thor",
                   "Peanut crusher",
                   "D.A.V.E",
                   "Andy Roid",
                   "Spanner mate",
                   "Drillbit 2000"]

    # Fill in form
    page = browser.page()
    page.select_option("#head", ROBOT_NAMES[int(order["Head"])-1] + " head")
    page.click("text=" + ROBOT_NAMES[int(order["Body"])-1]+ " body")
    page.locator("xpath=//label[contains(.,'3. Legs:')]/../input").fill(str(order["Legs"]))
    page.fill("#address", str(order["Address"]))

    # Submit form
    page.locator("#order").focus()
    page.locator("#order").click()

    # If order button is still visible we know an error has occured
    # Repeat until successful
    while page.locator("#order").is_visible(timeout=2000):
        page.locator("#order").focus()
        page.locator("#order").click()

def fill_form_with_excel_data():
    """
    Read data from excel and fill in orders
    """

    # Read CSV Input 
    tables = Tables()
    worksheet = tables.read_table_from_csv("orders.csv", header=True)

    # For each row in input...
    for row in worksheet:
        # Complete and submit the form
        fill_and_submit_form(row)

        # Store the receipt local output/receipts folder
        store_receipt_as_pdf(row["Order number"])

        # Store the screenshot local output/receipts folder
        screenshot_robot(row["Order number"])

        # Merge both the receipt and screenshot into one PDF
        embed_screenshot_to_receipt(row["Order number"])

        # Navigate back to start of loop
        reset_form()

def reset_form():
    """
    Get robot back into position to fill in another form after submitting successfully
    """
    # Click order another button and navigate back to form
    page = browser.page()
    page.locator("#order-another").click()

    close_annoying_modal()

def store_receipt_as_pdf(order_number):
    """
    Input: order_number, used to save receipt with unique and indentifiable path
    Export data to pdf
    """
    # Extract inner html of receipt element
    page = browser.page()
    order_result_html = page.locator("#receipt").inner_html()

    # Save html to local output/receipts file
    pdf = PDF()
    pdf.html_to_pdf(order_result_html, "output/receipts/" + order_number + ".pdf")

def screenshot_robot(order_number):
    """
    Input: order_number, used to save screenshot with unique and indentifiable path
    Take screenshot of the page and save to local output/receipts file
    Output screenshot path "output/receipts/screenshot[ORDERNUMBER].png"
    """
    page = browser.page()
    page.screenshot(path="output/receipts/screenshot" + order_number + ".png")

def embed_screenshot_to_receipt(order_number):
    """
    Input: order_number, used to match receipt to screenshot
    Merge an orders screenshot to the end of its matching receipt pdf file
    """
    pdf = PDF()

    pdf_path = "output/receipts/" + order_number + ".pdf"
    screenshot_path = "output/receipts/screenshot" + order_number + ".png"

    pdf.add_files_to_pdf([pdf_path, screenshot_path], target_document="output/full_receipt_" + order_number + ".pdf")

def archive_receipts():
    """
    Zip all files that being with the string "full_receipt" together
    """
    lib = Archive()
    lib.archive_folder_with_zip("./output", "receipts.zip", include="full_receipt*")