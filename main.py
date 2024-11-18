import ctypes
import datetime
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from pathlib import Path
import fitz
import openpyxl
import pymupdf
import requests
from aiohttp.web_routedef import static


class Berichtsheftmaker:
    def __init__(self):
        self.calenderweek = datetime.datetime.now().isocalendar()[1]
        self.currentyear = datetime.date.today().year
        self.stundenplan = f"StundenplanKW{self.calenderweek}"
        pdf: str = self.download_pdf()
        texfile: str = self.pdf_to_text(pdf)
        unedited_subjects: list = self.txt_to_list(texfile)
        cleaned_subjects: list = self.delete_dupe(unedited_subjects)
        self.listtoexcel(cleaned_subjects)
        self.send_mail()

    def download_pdf(self):
        url: str = f"https://service.viona24.com/stpusnl/daten/US_IT_2024_Winter_FIAE_B_2024_abKW{self.calenderweek}.pdf"
        filename: str = f"StundenplanKW{self.calenderweek}.pdf"
        filepath: Path = Path(filename)
        response = requests.get(url)
        filepath.write_bytes(response.content)
        print("success")
        return filename

    @staticmethod
    def pdf_to_text(file: str) -> str:
        doc = pymupdf.open(file)
        filename: str = f"output.txt"
        out = open(filename, "wb")
        for page in doc:
            text = page.get_text().encode("utf8")
            out.write(text)
            out.write(bytes((12,)))
        out.close()
        return filename

    @staticmethod
    def txt_to_list(data: str) -> list:
        with open(data, "r") as file:
            unedited_list: list = []
            for line in file:
                if "-" in line and ":" not in line:
                    fach = line.split(" ")[0]
                    unedited_list.append(fach)
                elif "16:00" in line:
                    unedited_list.append(".")
                elif "Mentor" in line and "Verf" in line:
                    unedited_list.append("Verfügungsstd.")
        return unedited_list

    @staticmethod
    def delete_dupe(data) -> list:
        edited_list: list = []
        for fach in range(len(data)):
            if data[fach] != data[fach - 1]:
                edited_list.append(data[fach])
            elif data[fach] == ".":
                edited_list.append(".")
        return edited_list

    def listtoexcel(self, data) -> None:
        os.remove(f"output.txt")
        os.remove(f"StundenplanKW{self.calenderweek}.pdf")
        path: str = "copycopy.xlsx"
        workbook = openpyxl.load_workbook(path)
        worksheet = workbook["Tabelle1"]
        counter: int = 4
        daycounter: int = 0
        praxiszeit: int = 420
        std_dauer: int = 90
        pausen_counter: int = 9
        praxis_counter: int = 8
        try:
            for fach in data:
                print(counter)
                if "Ver" not in fach and "." not in fach:
                    worksheet["B" + str(counter)] = str(fach) + ":"
                    worksheet["E" + str(counter)] = std_dauer
                    praxiszeit -= std_dauer
                    counter += 1
                elif "Ver" in fach:
                    worksheet["B" + str(counter)] = str(fach) + ":"
                    worksheet["E" + str(counter)] = std_dauer / 2
                    praxiszeit -= std_dauer / 2
                    counter += 1
                if "." in fach and "Ver" not in fach:
                    counter = 10 + (6 * daycounter)
                    daycounter += 1
                    worksheet["B" + str(praxis_counter)] = "Praxisunterricht:"
                    worksheet["E" + str(praxis_counter)] = praxiszeit
                    worksheet["E" + str(pausen_counter)] = 60
                    praxis_counter += 6
                    pausen_counter += 6
                    praxiszeit = 420
        except AttributeError:
            print(f"Error at {counter}")
        worksheet["D1"] = f"KW{self.calenderweek}"
        worksheet["E1"] = f"Jahr {self.currentyear}"
        workbook.save(f"Berichtsheft_KW{self.calenderweek}.xlsx")

    def send_mail(self) -> None:
        sender_email = os.getenv("SENDER_MAIL")
        password = os.getenv("GMAIL_PASSWORD")
        subject = "mail"
        body = "Body"
        recipient_email = os.getenv("RECIPIENT_MAIL")
        with open(f"Berichtsheft_KW{self.calenderweek}.xlsx", "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f"attachment; filename= Berichtsheft_KW{self.calenderweek}.xlsx")
        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = sender_email
        message['To'] = recipient_email
        html_part = MIMEText(body)
        message.attach(part)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, message.as_string())
        os.remove("Berichtsheft_KW47.xlsx")
        print("done")

App = Berichtsheftmaker

App()