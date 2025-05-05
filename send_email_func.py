import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import xlwings as xw
import streamlit as st

def send_forecast_email(MONTH_, DEADLINE_, SENDER_EMAIL_, SENDER_PASSWORD_):
    FILE_PATH = r"C:/Users/marselius.agus/Prospect Motor, PT/Marketing Support and Data Analyst - Documents/Target/2025/Publish/6. Target adj. Apr-Dec 2025.xlsx"
    tgt_table = pd.read_excel(FILE_PATH, sheet_name='RAW', engine='openpyxl', usecols='A:I', skiprows=1)

    MONTH_MAPPING = {
        'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
        'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
        'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
    }
    MONTH = MONTH_MAPPING.get(MONTH_, 0)
    STATUS = 'ADJ APR-DEC'

    mapping_model = {
        'ACCORD':'Accord', 'BRIO RS':'Brio RS',
        'BRIO SATYA': 'Brio Satya', 'CITY':'City', 
        'CITY HB':'City HB', 'CIVIC':'Civic', 
        'CIVIC TYPE R':'Civic Type R', 'MOBILIO':'Mobilio',
    }

    tgt_table['MODEL'] = tgt_table['MODEL'].replace(mapping_model)

    tgt_table = tgt_table[(tgt_table['STATUS'] == STATUS) & (tgt_table['BULAN'] == MONTH)].reset_index(drop=True)

    tgt_table['DEALER_NAME'] = tgt_table['DEALER'].str.replace('Honda ', '')

    mapping_month = {
        1:'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
        5:'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
        9:'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }

    tgt_table['MONTH'] = tgt_table['BULAN'].map(mapping_month)
    SELECTED_MONTH = tgt_table['MONTH'][0]

    master_dealer = pd.read_excel("data/master dealer.xlsx", sheet_name='Sheet')   

    tgt_table = tgt_table.merge(
        master_dealer,
        left_on="KD",
        right_on="DEALER ID",
        how="left"
    )

    tgt_table = tgt_table.drop(['DEALER ID', 'DEALER NAME'], axis=1)

    # Loop melalui setiap dealer
    for _, dealer_data in tgt_table.groupby('KD'):
        DEALER_CODE = dealer_data['KD'].iloc[0]
        DEALER_NICKNAME = dealer_data['DEALER NICK NAME'].iloc[0]

        # Generate file untuk dealer
        output_file = generate_forecast_file(dealer_data, SELECTED_MONTH, DEALER_NICKNAME)

        # Kirim email
        send_email_to_dealer(SENDER_EMAIL_, SENDER_PASSWORD_, DEALER_NICKNAME, SELECTED_MONTH, DEADLINE_, output_file)

    return "Email berhasil dikirim ke semua dealer!"

def generate_forecast_file(dealer_data, selected_month, dealer_nickname):
    app = xw.App(visible=False)
    wb = app.books.open(r'data/Format Forecast DO Dealer.xlsx')

    # Akses sheet
    sheet = wb.sheets['VERSI DEALER']

    # Modifikasi data
    sheet.range('D2').value = dealer_data['KD'].iloc[0]
    sheet.range('D3').value = dealer_nickname
    sheet.range('D22').value = dealer_nickname
    sheet.range('D41').value = dealer_nickname

    model_cell_map = {
        "Accord": 4,
        "Civic": 5,
        "Civic Type R": 7,
        "City": 8,
        "CR-V": 9,
        "WR-V": 10,
        "Brio RS": 11,
        "Brio Satya": 12,
        "Mobilio": 13,
        "HR-V 1.5": 15,
        "BR-V": 17,
        "City HB": 18
    }

    # Loop berdasarkan mapping
    for model, row in model_cell_map.items():
        qty = dealer_data[dealer_data['MODEL'] == model]['QTY'].sum()
        sheet.range(f'D{row}').value = qty
        
    NO_MONTH_MAPPING = {
        'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
        'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
        'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
    }
    number_month = NO_MONTH_MAPPING.get(selected_month, 0)
    
    # Simpan file
    output_file = f'D:/Otomatisasi Email/Forecast/{selected_month}/{number_month}. Forecast DO Dealer {selected_month} 2025 - {dealer_nickname}.xlsx'
    wb.save(output_file)
    wb.close()
    app.quit()

    return output_file

def send_email_to_dealer(sender_email, sender_password, dealer_nickname, selected_month, deadline, attachment_path):
    SMTP_SERVER = "smtp.office365.com"
    SMTP_PORT = 587



    RECIPIENT_EMAIL = "2173028@maranatha.ac.id" 
    CC_EMAIL = "agusmarselius@gmail.com"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = RECIPIENT_EMAIL
    msg['Cc'] = CC_EMAIL  
    msg['Subject'] = f"Forecast DO {selected_month} 2025 Dealer - {dealer_nickname}"

    HTML_BODY = f"""
    <html>
        <body>
            <p>Dear Team Dealer,</p>
            <p>
                Berikut ini kami lampirkan template data
                <b>Forecast DO pada bulan {selected_month} 2025</b> 
                yang harus diisikan oleh setiap dealer.
            </p>

            <p>
                Dimohon untuk dikirimkan kembali ke MD paling lambat hari 
                <b><span style="background-color: yellow;">{deadline}</span></b>
            </p>

            <p>
                Demikian informasi yang dapat kami sampaikan.
                Atas perhatiannya kami ucapkan terima kasih.    
            </p>
        </body>
    </html>
    """
    msg.attach(MIMEText(HTML_BODY, 'html'))

    # Attachment
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={attachment_path.split('/')[-1]}"
    )
    msg.attach(part)

    # Kirim email
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        recipients = [RECIPIENT_EMAIL] + [CC_EMAIL]
        server.sendmail(sender_email, recipients, msg.as_string())