import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
import base64
from send_email_func import send_forecast_email 
from market_info_func import send_marketing_info_email  

def horizontal_line():
    st.markdown(f'<hr>', unsafe_allow_html=True)
    
def enter():
    st.markdown('<br>', unsafe_allow_html=True)
    
def logo_link(path_img, width):
    st.markdown(
        """<div style="display: grid; place-items: center;">
        <img src="data:image/png;base64,{}" width="{}">
        </a></div>""".format(
            base64.b64encode(open(path_img, "rb").read()).decode(),
            width,
        ),
        unsafe_allow_html=True,
    )

FILE_PATH = r"C:/Users/marselius.agus/Prospect Motor, PT/Marketing Support and Data Analyst - Documents/Target/2025/Publish/6. Target adj. Apr-Dec 2025.xlsx"
data = pd.read_excel(FILE_PATH, sheet_name='RAW', engine='openpyxl', usecols='A:I', skiprows=1)

mapping_full_name_month = {
    1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
    5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
    9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
}

data['BULAN'] = data['BULAN'].map(mapping_full_name_month)

master_dealer = pd.read_excel("data/master dealer.xlsx", sheet_name='Sheet')   

data = data.merge(
    master_dealer,
    left_on="KD",
    right_on="DEALER ID",
    how="left"
)

data = data.drop(['DEALER ID', 'DEALER NAME'], axis=1)

st.set_page_config(
    page_title='Email Automation',
    page_icon=':mailbox_with_mail:',
    layout='wide'
)

with st.sidebar:
    st.markdown("""
        <div style='text-align: center; font-size:24px'>
            <b>
            Email Automation <br>
            Data Center Department
            </b>
        </div>
    """, unsafe_allow_html=True)
    
    enter()

    logo_link(r'Logo-PM.png', 150)
    horizontal_line()

    selected = option_menu(menu_title=None, 
                          options=['Forecast DO Dealer', 'Market Information', 'Claim Sales'], 
                        #   icons=['house'], 
                          menu_icon="cast", default_index=0
                        )

if selected == 'Forecast DO Dealer':
    st.markdown("""
        <div style='text-align: center; font-size:30px'>
            <b>Email Message Automation - Forecast DO Dealer</b>
        </div>
    """, unsafe_allow_html=True)  
    
    enter()
        
    list_bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
              'Agustus', 'September', 'Oktober', 'November', 'Desember']
    list_code_dealer = data['KD'].sort_values().unique()
        
    _, col_filter_bulan, col_filter_deadline, _ = st.columns([1,4,4,1]) 
    # with col_filter_code_dealer:
    #     SELECTED_CODE = st.selectbox('Dealer Code', list_code_dealer, index=0)

    with col_filter_bulan:
        SELECTED_MONTH = st.selectbox('Month', list_bulan, index=None)
    with col_filter_deadline:
        DEADLINE = st.text_input('Deadline', placeholder='Tanggal & Jam Deadline')
    
    _, col_input_email_sender, col_input_password_sender, _ = st.columns([1,4,4,1]) 
    
    with col_input_email_sender:
        SENDER_EMAIL = st.text_input('Email Sender')
    with col_input_password_sender:
        SENDER_PASSWORD = st.text_input('Email Password Sender', type='password')
    
    enter()

    _, col_display_email, _ = st.columns([1,8,1])
    
    
    with col_display_email:   
        # colored_header(
        #     label="Sample Tampilan Subject & Body Email yang Dikirim",
        #     description="",
        #     color_name="orange-70",
        # )
        horizontal_line()
        
        st.markdown("""
            <div style='text-align: center; font-size:20px'>
                <b>Sample Tampilan Subject & Body Email yang Dikirim</b>
            </div>
        """, unsafe_allow_html=True)

        horizontal_line()
    
        DEALER_NICKNAME = data[data['KD'] == 501]['DEALER NICK NAME'].drop_duplicates().values[0]
        
        if SELECTED_MONTH and DEADLINE:
            st.markdown(f"**Subject:** Forecast DO {SELECTED_MONTH} 2025 Dealer - {DEALER_NICKNAME}")
            st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")
            
            st.write("""
                <div style='text-align:left; font-size:16px'>
                    <p>Dear Team Dealer,</p>
                    <p>
                        Berikut ini kami lampirkan template data
                        <b>Forecast DO pada bulan {} 2025</b> 
                        yang harus diisikan oleh setiap dealer.
                    </p>
                    <p>
                        Dimohon untuk dikirimkan kembali ke MD paling lambat hari 
                        <b><span style="background-color: yellow;">{}</span></b>
                    </p>
                    <p>
                        Demikian informasi yang dapat kami sampaikan.
                        Atas perhatiannya kami ucapkan terima kasih.    
                    </p>
                </div>
            """.format(SELECTED_MONTH, DEADLINE), unsafe_allow_html=True)
        else:
            st.warning("Harap lengkapi semua input (Month dan Deadline) sebelum melihat preview email.")    
        
        horizontal_line()

    _, col_button_success,_ = st.columns([2,2,2])

    with col_button_success:
        if st.button("Blast Email", use_container_width=True):
            # Kirim email ke semua dealer
            outcome = send_forecast_email(SELECTED_MONTH, DEADLINE, SENDER_EMAIL, SENDER_PASSWORD)
            st.success(outcome, icon="✅")
                        
if selected == 'Market Information':
    st.markdown("""
        <div style='text-align: center; font-size:30px'>
            <b>Email Message Automation - Market Information</b>
        </div>
    """, unsafe_allow_html=True)  
    
    enter()
    
    provinsi_mapping = {
        'D.I. Aceh':'Aceh', 'North Sumatera':'Sumatera Utara', 'West Sumatera':'Sumatera Barat', 'Batam':'Batam',
        'Kep. Riau':'Kepulauan Riau', 'Riau':'Riau', 'Jambi':'Jambi', 'Bangka-Belitung':'Bangka Belitung', 'South Sumatera':'Sumatera Selatan',
        'Bengkulu':'Bengkulu', 'Lampung':'Lampung', 'Cikarang':'Cikarang', 'South Kalimantan':'Kalimantan Selatan',
        'Central Kalimantan':'Kalimantan Tengah', 'East Kalimantan':'Kalimantan Timur', 'West Kalimantan':'Kalimantan Barat',
        'South Sulawesi':'Sulawesi Selatan', 'Central Sulawesi':'Sulawesi Tengah', 'North Sulawesi':'Sulawesi Utara',
        'Southeast Sulawesi':'Sulawesi Tenggara', 'Gorontalo':'Gorontalo', 'West Sulawesi':'Sulawesi Barat', 'Maluku':'Maluku',
        'North Maluku':'Maluku Utara', 'Papua':'Papua'
    }
    
    list_bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
              'Agustus', 'September', 'Oktober', 'November', 'Desember']
    
    data['PROVINCE'] = data['PROVINCE'].map(provinsi_mapping)
    
    _, col_filter_area, col_filter_bulan, col_link_power_bi, _ = st.columns([1,3,2,3,1]) 
    with col_filter_area:
        # SELECTED_CODE = st.multiselect('Select Dealer Code', list_area)
        SELECTED_AREA = st.selectbox('Area', data['PROVINCE'].unique(), index=None)
    with col_filter_bulan:
        PERIOD_MONTH = st.text_input('Period', placeholder='Ex: Jan-Feb')
    with col_link_power_bi:
        LINK_POWER_BI = st.text_input('Link Power BI', placeholder='Input Here')
    
    _, col_input_email_sender, col_input_password_sender, _ = st.columns([1,4,4,1]) 
    
    with col_input_email_sender:
        SENDER_EMAIL = st.text_input('Email Sender')
    with col_input_password_sender:
        SENDER_PASSWORD = st.text_input('Email Password Sender', type='password')
    
    enter()

    _, col_display_email, col_filter_versi_email, _ = st.columns([1,8,2,1])
    
    with col_filter_versi_email:
        VERSI_EMAIL = st.selectbox('Versi Email', ['Versi HPM', 'Versi DLR', 'Versi Polreg Dealer'], index=None)
    
    with col_display_email:   
        horizontal_line()

        st.markdown("""
            <div style='text-align: center; font-size:20px'>
                <b>Sample Tampilan Subject & Body Email yang Dikirim</b>
            </div>
        """, unsafe_allow_html=True)

        horizontal_line()
                
        if PERIOD_MONTH and SELECTED_AREA and LINK_POWER_BI and VERSI_EMAIL:
            if VERSI_EMAIL == 'Versi HPM': 
                st.markdown(f"**Subject:** Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025")
                st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")
                
                st.html(f"""
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
                st.markdown(f"**Subject:** Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 DLR Version")
                st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")

                st.html(f"""
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

                        <p>Link: <a href="{LINK_POWER_BI}" style="color:blue">Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025</a></p>

                        <p>
                            Atas perhatian bapak/ibu, kami ucapkan terima kasih. <br>
                            Regards,
                        </p>
                    </div>
                """)

            if VERSI_EMAIL == 'Versi Polreg Dealer': 
                st.markdown(f"**Subject:** Power BI Market Information {SELECTED_AREA} {PERIOD_MONTH} 2025 Versi Polreg Dealer")
                st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")

                st.html(f"""
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
                
        else:
            st.warning("Harap lengkapi semua input (Area, Periode, Link Power BI, dan Versi Email) sebelum melihat preview email.")    
                
    _, h_line, _ = st.columns([1,8,1])
    
    with h_line:
        horizontal_line()
    
    _, col_button_success,_ = st.columns([2,2,2])

    with col_button_success:
        if st.button("Send Email", use_container_width=True):
            outcome = send_marketing_info_email(SELECTED_AREA, PERIOD_MONTH, LINK_POWER_BI, VERSI_EMAIL, SENDER_EMAIL, SENDER_PASSWORD)
            st.success(outcome, icon="✅")
            
if selected == 'Claim Sales':
    st.markdown("""
        <div style='text-align: center; font-size:30px'>
            <b>Email Message Automation - Claim Sales</b>
        </div>
    """, unsafe_allow_html=True)  
    
    enter()
    
    list_bulan = ['Januari', 'Februari', 'Maret']
    list_code_dealer = [501, 502, 503]
    
    _, col_filter_code_dealer, col_filter_bulan, col_filter_deadline, _ = st.columns([1,3,2,3,1]) 
    with col_filter_code_dealer:
        # SELECTED_CODE = st.multiselect('Select Dealer Code', list_code_dealer)
        SELECTED_CODE = st.selectbox('Dealer Code', list_code_dealer, index=None)
    with col_filter_bulan:
        SELECTED_MONTH = st.selectbox('Month', list_bulan, index=None)
    with col_filter_deadline:
        DEADLINE = st.text_input('Deadline', placeholder='YYYY-MM-DD')
    
    _, col_input_email_sender, col_input_password_sender, _ = st.columns([1,4,4,1]) 
    
    with col_input_email_sender:
        sender_email = st.text_input('Email Sender')
    with col_input_password_sender:
        sender_password = st.text_input('Email Password Sender', type='password')
    
    enter()

    _, col_display_email, _ = st.columns([1,8,1])
    
    with col_display_email:   
        horizontal_line()

        st.markdown("""
            <div style='text-align: center; font-size:20px'>
                <b>Sample Tampilan Subject & Body Email yang Dikirim</b>
            </div>
        """, unsafe_allow_html=True)
        
        horizontal_line()
        
        if SELECTED_MONTH and SELECTED_CODE and DEADLINE:
            st.markdown(f"**Subject:** Forecast DO {SELECTED_MONTH} 2025 Dealer - {SELECTED_CODE}")
            st.markdown("**To:** [Recipient Email] | **CC:** [Email CC]")
            
            st.write("""
                <div style='text-align:left; font-size:16px'>
                    <p>Dear Team Dealer,</p>
                    <p>
                        Berikut ini kami lampirkan template data
                        <b>Forecast DO pada bulan {}</b> 
                        yang harus diisikan oleh setiap dealer.
                    </p>
                    <p>
                        Dimohon untuk dikirimkan kembali ke MD (Main Dealer) paling lambat hari 
                        <b><span style="background-color: yellow;">{}</span></b>
                    </p>
                    <p>
                        Demikian informasi yang dapat kami sampaikan.
                        Atas perhatiannya kami ucapkan terima kasih.    
                    </p>
                </div>
            """.format(SELECTED_MONTH, DEADLINE), unsafe_allow_html=True)
        else:
            st.warning("Harap lengkapi semua input (Dealer Code, Month, dan Deadline) sebelum melihat preview email.")    
        
        horizontal_line()
        
        uploaded_files = st.file_uploader(
            "Upload file here", accept_multiple_files=True
        )
        for uploaded_file in uploaded_files:
            # bytes_data = uploaded_file.read()
            st.write("Filename:", uploaded_file.name)
            # st.write(bytes_data)
    
    enter()

    _, col_button_success,_ = st.columns([2,2,2])

    with col_button_success:
        if st.button("Send Email", use_container_width=True):
            st.success('Email Sucessfully Sent', icon="✅")


        