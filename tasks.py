from http.client import CannotSendRequest
import os
from pathlib import Path
from socket import timeout
import requests
from robocorp import browser
from robocorp.browser import Page
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

FILE_NAME = "orders.csv"
CSV_URL = f"https://robotsparebinindustries.com/{FILE_NAME}"
OUTPUT_DIR = Path(os.getenv("ROBOT_ARTIFACTS", "output"))
CURDIR = Path(__file__).parent
ORDERS_DIR = OUTPUT_DIR / "orders"
ORDERS_DIR.mkdir(parents=True, exist_ok=True)


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    orders = get_orders()
    page = open_robot_order_website()
    page.set_default_timeout(10000)
    go_to_order_page(page)
    for order in orders:
        close_annoying_modal(page)
        fill_the_form(page, order)
        submit_order(page)
        while page.is_visible(".alert-danger"):
            submit_order(page)
        store_receipt_as_pdf(page, order["Order number"])
        create_new_order(page)
    archive_receipts()


def create_new_order(page):
    page.get_by_role("button", name="Order another robot").click()


def store_receipt_as_pdf(page: browser.Page, order_number: int) -> Path:
    """
    Stores a receipt as a PDF file by capturing screenshots of the receipt and a robot image.
    This function generates a PDF document containing the receipt and associated images for a specified order number.

    Args:
        page (browser.Page): The page object from which to capture the receipt screenshot.
        order_number (int): The order number used to name the output files.

    Returns:
        Path: The path to the generated PDF file.

    Raises:
        ValueError: If the receipt element is not found on the page.
        FileNotFoundError: If the screenshot of the receipt is not saved correctly.
    """

    pdf = PDF()
    pdf_file = ORDERS_DIR / f"order_{order_number}.pdf"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    receipt_file = OUTPUT_DIR / f"receipt_{order_number}.png"
    robot_file = OUTPUT_DIR / f"robot_{order_number}.png"

    if receipt_element := page.locator("#receipt"):
        receipt_element.screenshot(path=str(receipt_file))
        screenshot_robot(page, robot_file)
    else:
        raise ValueError("Elemento '#receipt' não encontrado na página.")

    list_of_files = [str(receipt_file), str(robot_file)]
    if receipt_file.is_file():
        pdf.add_files_to_pdf(files=list_of_files, target_document=str(pdf_file))
    else:
        raise FileNotFoundError(
            f"Screenshot não foi salvo corretamente em {receipt_file}"
        )

    return pdf_file


def screenshot_robot(page, robot_file):
    return page.locator("#robot-preview-image").screenshot(path=str(robot_file))


def open_robot_order_website():
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
        slowmo=0,
    )
    return browser.goto("https://robotsparebinindustries.com/")


def go_to_order_page(page):
    page.get_by_role("link", name="Order your robot!").click()


def close_annoying_modal(page):
    if page.is_visible(".modal-header"):
        page.get_by_role("button", name="OK").click()


def fill_the_form(page: Page, order):
    """
    Fills out a form on a given page using the details from an order.
    This function sets the head, body, legs, and address fields of the form based on the provided order data.

    Args:
        page (Page): The page object where the form is located.
        order (dict): A dictionary containing order details, including Head, Body, Legs, and Address.

    Returns:
        None
    """

    set_head(page, order["Head"])
    set_body(page, order["Body"])
    set_legs(page, order["Legs"])
    set_address(page, order["Address"])


def submit_order(page: Page):
    page.get_by_role("button", name="Order").click()


def set_address(page: Page, address):
    page.get_by_placeholder("Shipping address").fill(address)


def set_legs(page: Page, legs):
    page.get_by_placeholder("Enter the part number").fill(legs)


def set_head(page: Page, head):
    page.select_option("select#head.custom-select", head)


def set_body(page: Page, body):
    page.click(f"#id-body-{body}")


def get_orders():
    """
    Retrieves order data by downloading a CSV file and reading its contents into a structured format.
    This function specifically extracts relevant columns from the CSV to provide a clear view of the orders.

    Args:
        None

    Returns:
        list: A list of dictionaries containing order details, including order number, head, body, legs, and address.
    """

    csv_file = download_file(CSV_URL, target_dir=CURDIR, target_filename=FILE_NAME)
    tables = Tables()
    return tables.read_table_from_csv(
        csv_file, columns=["Order number", "Head", "Body", "Legs", "Address"]
    )


def archive_receipts():
    """
    Archives receipts by compressing the contents of the specified orders folder into a ZIP file.
    This function utilizes the Archive library to perform the archiving operation.

    Args:
        None

    Returns:
        None
    """
    lib = Archive()
    lib.archive_folder_with_zip(
        "./output/orders/", f"{OUTPUT_DIR}/receipts.zip", recursive=True
    )


def download_file(url: str, *, target_dir: Path, target_filename: str) -> Path:
    """
    Downloads a file from the given URL into a custom folder & name.

    Args:
        url: The target URL from which we'll download the file.
        target_dir: The destination directory in which we'll place the file.
        target_filename: The local file name inside which the content gets saved.

    Returns:
        Path: A Path object pointing to the downloaded file.
    """
    # Obtain the content of the file hosted online.
    response = requests.get(url)
    response.raise_for_status()  # this will raise an exception if the request fails
    # Write the content of the request response to the target file.
    target_dir.mkdir(exist_ok=True)
    local_file = target_dir / target_filename
    local_file.write_bytes(response.content)
    return local_file
