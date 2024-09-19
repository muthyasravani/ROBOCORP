from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def Order_CreationBot ():
    """It creates the Order in the Website"""
    """Download the Receipts """
    """Take the screenshot of the Ordered Robot"""
    """"Append the Robot to the Receipt PDF"""
    """"Archive the PDF into Zip folder"""
    
    browser.configure(slowmo=800)
    Open_RobotOrder_Website()
    Download_OrdersFile()
    get_orders()
    archive_receipts()
    

def Open_RobotOrder_Website():

    """"Open the Order Creation website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    page=browser.page()
    page.click('text=OK')


def Download_OrdersFile():

    """"Downlaod the orders File"""
    http=HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",overwrite=True)

def get_orders():

    """"Get the orders CSV data into Table"""
    csvfile=Tables()
    Orderdata= csvfile.read_table_from_csv("orders.csv")
    for row in Orderdata:
        OrderRobot(row)

def OrderRobot(OrderDetails):
    """"Fill the Form with robot order details"""

    page=browser.page()
    page.select_option("#head",OrderDetails["Head"])
    page.check(f"#id-body-"+OrderDetails["Body"])
    page.fill(".form-control",OrderDetails["Legs"])
    page.fill("#address",OrderDetails["Address"])
    page.click("#preview")
    page.click("#order")
    while not page.query_selector("#order-another"):
        page.click("#order")
    CreateReceipt(OrderDetails["Order number"])
    page.click("#order-another")
    page.click('text=OK')

def CreateReceipt(ordernumber):
    """Create the receipt as PDF and take the screenshot of the"""
    page=browser.page()
    orderreceipt= page.locator("#receipt").inner_html()
    pdf = PDF()
    pdfPath="output/{}.pdf".format(ordernumber)
    pdf.html_to_pdf(orderreceipt,pdfPath)

    robo = page.query_selector("#robot-preview-image")
    screenshot = f"output/{ordernumber}.png"
    robo.screenshot(path=screenshot)
    MergeImageandPDF(pdfPath,screenshot)

def MergeImageandPDF(pdfpath,screenshot):
    """"Append the Robot Image to the PDF of the receipt"""
    pdf=PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdfpath, append=True)

def archive_receipts():
    """Archives the receipts""" 
    folder = Archive()
    folder.archive_folder_with_zip('output', 'output/orders.zip', include='*.pdf')










   

    



















