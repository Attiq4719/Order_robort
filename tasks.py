from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Archive import Archive
import time
import logging
import csv

# creating the logger object
logger = logging.getLogger()

class order_robots_from_RobotSpareBin:
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    
    def __init__(self):
        self.browser = Selenium()
        self.http = HTTP()
        self.excel = Files()
        self.pdf = PDF()
        self.archive = Archive()

    def open_robot_order_website(self):
        """Opens the browser with the given URL"""
        self.browser.open_available_browser("https://robotsparebinindustries.com/#/robot-order",maximized=True)

    def download_csv_file(self):
        """Downloads excel file from the given URL"""
        self.http.download(url="https://robotsparebinindustries.com/orders.csv", target_file="output/",overwrite=True)


    def handle_modal(self):
        """Handles the modal dialog for accepting cookies"""
        time.sleep(1)
        self.browser.click_button("//button[@class = 'btn btn-dark']")

    def get_order(self):
        """Read data from orders.csv and return a list of dictionaries"""
        with open('output/orders.csv', mode='r', newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            orders = [row for row in reader]
        return orders

    def fill_the_form(self, order):
        """Fills in the form with the given order data"""
        self.browser.select_from_list_by_value("//select[@id='head']", order["Head"])
        self.browser.find_element(f'//div[@class="stacked"]//label[@for="id-body-{order["Body"]}"]').click()
        self.browser.input_text("//input[@placeholder='Enter the part number for the legs']", order["Legs"])
        self.browser.input_text("//input[@id='address']", order["Address"])
        self.browser.capture_page_screenshot(f"output/before_scrol_screenshot_{order['Order number']}.png")
        if self.browser.does_page_contain_element("//button[@id='order']"):
            # element = self.browser.find_element("//button[@id='order']")
            self.browser.execute_javascript("window.scrollTo(0, document.body.scrollHeight);")
            # self.browser.execute_javascript("document.querySelector('#order').scrollIntoView(true);")
        self.browser.click_button("//button[@id='order']")
        # Check if the order was successful
        while self.browser.is_element_visible("//div[@class='alert alert-danger']"):
            self.browser.capture_page_screenshot(f"output/before_scroll_screenshot_{order['Order number']}.png")
            if self.browser.does_page_contain_element("//button[@id='order']"):
                self.browser.execute_javascript("window.scrollTo(0, document.body.scrollHeight);")
                # element = self.browser.find_element("//button[@id='order']")
                # self.browser.execute_javascript("document.querySelector('#order').scrollIntoView(true);")
            self.browser.click_button("//button[@id='order']")

    def store_receipt_as_pdf(self,order_number):
        """Store the receipt as a PDF file"""
        # Get the HTML content of the receipt
        receipt_html = self.browser.get_element_attribute("//div[@id='order-completion']", "outerHTML")
        # Save the HTML content as a PDF file
        self.pdf.html_to_pdf(receipt_html, f"output/receipt_{order_number}.pdf")
        self.browser.click_button("//button[@id='order-another']")
        return f"output/receipt_{order_number}.pdf"
    
    def screenshot_robot(self,order_number):
        # Wait for the receipt to be generated
        time.sleep(2)
        # Take a screenshot of the robot
        s1 = self.browser.capture_element_screenshot("//div[@id='robot-preview-image']",f"output/robort_{order_number}.png")
        return f"output/robort_{order_number}.png"
    
    def embed_screenshot_to_pdf(self,snap_address,pdf_address):
        """Embed the screenshot into the PDF file"""
        # Open the PDF file
        self.pdf.open_pdf(pdf_address)
        # Add the screenshot to the PDF file
        self.pdf.add_watermark_image_to_pdf(image_path=snap_address,source_path=pdf_address,output_path=pdf_address)
        # self.pdf.add_files_to_pdf(snap_address,pdf_address,append=True)
        # Save the PDF file
        self.pdf.save_pdf(pdf_address)
    
    def archive_receipts(self):
        """Archive the receipts and images"""
        # Create a ZIP archive of the receipts and images
        self.archive.archive_folder_with_zip("output","output/receipt_pdf.zip",include="*.pdf")
        self.archive.archive_folder_with_zip("output","output/robot_images.zip",include="*.png")

    def close_browser(self):
        """Closes the browser"""
        self.browser.close_browser()

@task
def order_robot():
    order_robot_spare_bin = order_robots_from_RobotSpareBin()
    logging.info("Download csv data...")
    order_robot_spare_bin.download_csv_file()
    logging.info("Opened browser...")
    orders = order_robot_spare_bin.get_order()
    logging.info("Get order form csv...")
    order_robot_spare_bin.open_robot_order_website()
    for order in orders:
        print(order)
        order_robot_spare_bin.handle_modal()
        logging.info("Handle Modal...")
        order_robot_spare_bin.fill_the_form(order)
        logging.info("Fill form with csv data...") 
        print("order plased susfully")
        ordernumber = order["Order number"]
        snap_address = order_robot_spare_bin.screenshot_robot(ordernumber)
        logging.info("Screenshot of the robot...")
        pdf_address = order_robot_spare_bin.store_receipt_as_pdf(ordernumber)
        logging.info("Store receipt as pdf...")
        order_robot_spare_bin.embed_screenshot_to_pdf(snap_address,pdf_address)
        logging.info("Embed screenshot to pdf...")

    order_robot_spare_bin.archive_receipts()
    logging.info("Archive receipts...")
    order_robot_spare_bin.close_browser()
    logging.info("Closing browser...")
    logging.info("All tasks completed successfully.") 