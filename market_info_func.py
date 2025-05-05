import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import xlwings as xw
import streamlit as st

def send_marketing_info_email(AREA_, PERIOD_, LINK_POWERBI_, VERSI_EMAIL_, SENDER_EMAIL_, SENDER_PASSWORD_):
    
    SELECTED_AREA = AREA_
    PERIOD_MONTH = PERIOD_
    LINK_POWER_BI = LINK_POWERBI_
    VERSI_EMAIL = VERSI_EMAIL_
    SENDER_EMAIL = SENDER_EMAIL_
    SENDER_PASSWORD = SENDER_PASSWORD_
    
    SMTP_SERVER = "smtp.office365.com"       ## Server SMTP Outlook
    SMTP_PORT = 587
            
    RECIPIENT_EMAIL = 'thismad07@gmail.com'
    CC_EMAIL = 'agusmarselius@gmail.com'
    
    ## Email
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECIPIENT_EMAIL
    msg['Cc'] = CC_EMAIL  

    if VERSI_EMAIL == 'Versi HPM': 
        msg['Subject'] = f"Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025"
        
        HTML_BODY = (f"""
            <div style='text-align:left; font-size:16px'>
                <p>
                    Kepada Yth, <br>
                    Bapak/Ibu <br>
                    Pimpinan Dealer Honda 
                <p>

                <p>
                    Berikut ini kami sampaikan link Power BI untuk informasi market {SELECTED_AREA} periode {PERIOD_MONTH} 2025. <br>
                    Data market berdasarkan Data Polreg yang HPM kirimkan kepada Main Dealer. <br>
                    Untuk keterangan lebih lanjut mengenai data market ini, silakan untuk menghubungi kami kembali.
                </p>

                <p>Link: <a href="{LINK_POWER_BI}" style="color:blue" target="_blank">Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025</a></p>

                <p>
                    Atas perhatian bapak/ibu, kami ucapkan terima kasih. <br>
                    Regards,
                </p>
            </div>
        """)


    if VERSI_EMAIL == 'Versi DLR': 
        msg['Subject'] = f"Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 DLR Version"
        # st.markdown(f"**Subject:** Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 DLR Version")
        # st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")

        HTML_BODY = (f"""
            <div style='text-align:left; font-size:16px'>
                <p>
                    Kepada Yth, <br>
                    Bapak/Ibu <br>
                    Pimpinan Dealer Honda        
                </p>

                <p>
                    Berikut ini kami sampaikan link Power BI untuk informasi market {SELECTED_AREA} periode {PERIOD_MONTH} 2025. <br>
                    Data market berdasarkan Data Polreg yang Dealer kirimkan kepada Main Dealer. <br>
                    Untuk keterangan lebih lanjut mengenai data market ini, silakan untuk menghubungi kami kembali.
                </p>

                <p>Link: <a href="{LINK_POWER_BI}" style="color:blue">Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025</a></p

                <p>
                    Atas perhatian bapak/ibu, kami ucapkan terima kasih. <br>
                    Regards,
                </p>
            </div>
        """)

    if VERSI_EMAIL == 'Versi Polreg Dealer': 
        msg['Subject'] = f"Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 Versi Polreg Dealer"
        
        # st.markdown(f"**Subject:** Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 Versi Polreg Dealer")
        # st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")

        HTML_BODY = (f"""
            <div style='text-align:left; font-size:16px'>
                <p>
                    Kepada Yth, <br>
                    Bapak/Ibu <br>
                    Pimpinan Dealer Honda        
                </p>

                <p>
                    Berikut ini kami sampaikan link Power BI untuk informasi market {SELECTED_AREA} periode {PERIOD_MONTH} 2025. <br>
                    Data total Honda {SELECTED_AREA} menggunakan <b>data Notice STNK Dealer</b> dan data competitor menggunakan <b>data Polreg Dealer</b>. <br>
                    Untuk keterangan lebih lanjut mengenai data market ini, silakan untuk menghubungi kami kembali.
                </p>

                <p>Link: <a href="{LINK_POWER_BI}" style="color:blue">Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025</a></p>

                <p>
                    Atas perhatian bapak/ibu, kami ucapkan terima kasih. <br>
                    Regards,
                </p>
            </div>
        """)
        
    msg.attach(MIMEText(HTML_BODY, 'html'))  ## Menambahkan body email dalam format HTML

    ## Kirim email melalui server SMTP
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  ## Inisialisasi koneksi TLS
            server.login(SENDER_EMAIL, SENDER_PASSWORD)  ## Login ke server SMTP
            recipients = [RECIPIENT_EMAIL] + [CC_EMAIL]
            server.sendmail(SENDER_EMAIL, recipients, msg.as_string())  ## Kirim email
            print("Email berhasil dikirim!")
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

    return "Email berhasil dikirim!"