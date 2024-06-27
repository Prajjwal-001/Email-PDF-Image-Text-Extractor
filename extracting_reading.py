import imaplib
import email
import os
import fitz
from PIL import Image
from pytesseract import pytesseract
import enum
import logging
import configparser

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

config = configparser.ConfigParser()
config.read('config.ini')

ORG_EMAIL = "@gmail.com"
useName = config['EMAIL']['User'] + ORG_EMAIL
passWord = config['EMAIL']['Password']

imap_url = 'imap.gmail.com'


class OS(enum.Enum):
    Mac = 0
    Windows = 1

class Language(enum.Enum):
    ENG = 'eng'
    RUS = 'rus'
    ITA = 'ita'
    HIN = 'hin'
    ENG_RUS = 'eng+rus'
    ENG_HIN = 'eng+hin'


def fetch_email_attachments():
    try:
        my_mail = imaplib.IMAP4_SSL(imap_url)
        my_mail.login(useName, passWord)
        my_mail.select('Inbox')

        data = my_mail.search(None, 'All')
        mail_ids = data[1]
        id_list = mail_ids[0].split(b' ')
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        email_count = 0

        for i in range(latest_email_id, first_email_id, -1):
            if email_count >= 15:
                break

            data = my_mail.fetch(str(i), '(RFC822)')
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_bytes(arr[1])
                    email_subject = msg['subject']
                    index = (msg['from']).find('<')
                    email_from = (msg['from'])[0:index]
                    email_date = msg['Date']

                    logging.info(f'From: {email_from}')
                    logging.info(f'Subject: {email_subject}')
                    logging.info(f'Date: {email_date}')

                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue
                        fileName = part.get_filename()
                        if bool(fileName):
                            filePath = os.path.join(config['PATHS']['Attachments'], fileName)
                            if not os.path.isfile(filePath):
                                with open(filePath, 'wb') as fp:
                                    fp.write(part.get_payload(decode=True))
            email_count += 1
        my_mail.close()
        my_mail.logout()
    except Exception as e:
        logging.error(f"An error occurred while fetching emails: {e}")


def extract_images_from_pdfs(pdf_folder, output_folder):
    try:
        pdf_files = [os.path.join(pdf_folder, file) for file in os.listdir(pdf_folder) if file.endswith('.pdf')]

        for pdf_file in pdf_files:
            extract_images_from_pdf(pdf_file, output_folder)
    except Exception as e:
        logging.error(f"An error occurred while extracting images from PDFs: {e}")

def extract_images_from_pdf(file_path, output_folder):
    try:
        pdf = fitz.open(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]

        os.makedirs(output_folder, exist_ok=True)

        for page_num in range(len(pdf)):
            page = pdf.load_page(page_num)
            image_list = page.get_images(full=True)

            if image_list:
                logging.info(f"Found {len(image_list)} images on page {page_num + 1} of {file_path}")

                for img_index, img_info in enumerate(image_list, start=1):
                    xref = img_info[0]  # retrieves the XRef number of the image
                    base_image = pdf.extract_image(xref)

                    image_name = f"{base_name}_page{page_num + 1}_img{img_index}.png"
                    image_path = os.path.join(output_folder, image_name)

                    with open(image_path, "wb") as f:
                        f.write(base_image["image"])
            else:
                logging.info(f"No images found on page {page_num + 1} of {file_path}")

        pdf.close()
    except Exception as e:
        logging.error(f"An error occurred while extracting images from {file_path}: {e}")

class ImageReader:
    def __init__(self, os_type: OS):
        if os_type == OS.Windows:
            windows_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            pytesseract.tesseract_cmd = windows_path
            logging.info("Running on Windows")
        elif os_type == OS.Mac:
            mac_path = r'/usr/local/bin/tesseract'
            pytesseract.tesseract_cmd = mac_path
            logging.info("Running on Mac")
    
    def extract_text(self, image: str, lang: Language) -> str:
        img = Image.open(image)
        extracted_text = pytesseract.image_to_string(img, lang=lang.value)
        return extracted_text
    
    def extract_text_from_images(self, images_folder: str, lang: Language):
        image_files = [os.path.join(images_folder, file) for file in os.listdir(images_folder) if file.endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]
        
        all_texts = []
        for image_file in image_files:
            logging.info(f"Extracting text from {image_file}...")
            text = self.extract_text(image_file, lang)
            all_texts.append((image_file, text))
        
        return all_texts

def main():
    fetch_email_attachments()
    
    pdf_folder = config['PATHS']['Attachments']
    output_folder = config['PATHS']['ExtractedImages']

    extract_images_from_pdfs(pdf_folder, output_folder)

    ir = ImageReader(OS.Windows)
    extracted_texts = ir.extract_text_from_images(output_folder, lang=Language.ENG)
    
    for image_file, text in extracted_texts:
        logging.info(f"\nText from {image_file}:\n{text}\n")

if __name__ == '__main__':
    main()
