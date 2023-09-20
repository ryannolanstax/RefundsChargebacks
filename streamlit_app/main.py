import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import json
import numpy as np
import datetime
from datetime import date, timedelta
import io

def download_button(objects_to_download, download_filename):
    """
    Generates a link to download the given objects_to_download as separate sheets in an Excel file.
    Params:
    ------
    objects_to_download (dict): A dictionary where keys are sheet names and values are objects to be downloaded.
    download_filename (str): filename and extension of the Excel file. e.g. mydata.xlsx
    Returns:
    -------
    (str): the anchor tag to download the Excel file with multiple sheets
    """
    # Create an in-memory Excel writer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as excel_writer:
        for sheet_name, object_to_download in objects_to_download.items():
            if isinstance(object_to_download, pd.DataFrame):
                # Write DataFrame as a sheet
                object_to_download.to_excel(excel_writer, sheet_name=sheet_name, index=False)
            else:
                # Convert other objects to a DataFrame and write as a sheet
                df = pd.DataFrame({"Data": [object_to_download]})
                df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    # Seek to the beginning of the in-memory stream
    output.seek(0)
    excel_data = output.read()

    # Encode the Excel file to base64 for download
    b64 = base64.b64encode(excel_data).decode()

    dl_link = f"""
    <html>
    <head>
    <title>Start Auto Download file</title>
    <script src="http://code.jquery.com/jquery-3.2.1.min.js"></script>
    <script>
    $('<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{download_filename}">')[0].click()
    </script>
    </head>
    </html>
    """
    return dl_link

def download_df():
    if uploaded_files:
        for file in uploaded_files:
            file.seek(0)
        uploaded_data_read = [pd.read_csv(file) for file in uploaded_files]
        df = pd.concat(uploaded_data_read)

        days90 = date.today() - timedelta(days=90)
        days180 = date.today() - timedelta(days=180)

        df2 = df[df['created_at'].str.split(expand=True)[1].isna() == False]
        dfbaddates = df[df['created_at'].str.split(expand=True)[1].isna() == True].copy()

        dfbaddates['created_at'] = dfbaddates['created_at'].apply(lambda x: datetime.datetime(1900, 1, 1, 0, 0, 0) + datetime.timedelta(days=float(x)))
        dfbaddates['created_at'] = dfbaddates['created_at'].dt.strftime('%m/%d/%y')
        df.loc[df['created_at'].str.split(expand=True)[1].isna() == False, 'created_at'] = df2['created_at'].str.split(expand=True)[0].str.strip()
        newdf = pd.concat([df2, dfbaddates])
        newdf['created_at'] = pd.to_datetime(newdf['created_at']).dt.date
        df3 = newdf.query("payment_method == 'card' | payment_method == 'bank'")
        df3.drop(['id','merchant_id','user_id','customer_id','subtotal','tax','is_manual','success','donation','tip','meta','pre_auth','updated_at','source', 'issuer_auth_code'], axis=1, inplace=True)
        df4 = df3.loc[:,['type', 'created_at', 'total', 'payment_person_name', 'customer_firstname', 'customer_lastname',\
            'payment_last_four', 'last_four', 'payment_method', 'channel', 'memo', 'payment_note', 'reference', \
            'payment_card_type', 'payment_card_exp', 'payment_bank_name', 'payment_bank_type',\
            'payment_bank_holder_type', 'billing_address_1', 'billing_address_2','billing_address_city', \
            'billing_address_state', 'billing_address_zip', 'customer_company','customer_email', 'customer_phone', \
            'customer_address_1','customer_address_2', 'customer_address_city', 'customer_address_state', \
            'customer_address_zip', 'customer_notes', 'customer_reference', 'customer_created_at', \
            'customer_updated_at', 'customer_deleted_at', 'gateway_id', 'gateway_name', 'gateway_type', \
            'gateway_created_at', 'gateway_deleted_at', 'user_name', 'system_admin', 'user_created_at',\
            'user_updated_at', 'user_deleted_at']]

        volume = df3.query("type == 'charge'")
        volumetotal = np.sum(volume['total'])
        volume90 = np.sum(volume[volume['created_at'] > days90]['total'])
        volume180 = np.sum(volume[volume['created_at'] > days180]['total'])

        refund = df3.query("type == 'refund'")
        refundtotal = np.sum(refund['total'])
        refund90 = np.sum(refund[refund['created_at'] > days90]['total'])

        Lifetime_refund_rate = (refundtotal / volumetotal)
        day90_refund_rate = (refund90/volume90)
        Lifetime_chargeback_rate = (chargebackslifetime / volumetotal)
        day180_chargeback_rate = (chargebacks180/volume180)

        dfcalc = pd.DataFrame({'Refunds for past 90 days':[refund90],
                        '90 day volume':[volume90],
                        '90 day refund rate':[day90_refund_rate],
                        'Lifetime refunds':[refundtotal],
                        'Lifetime volume':[volumetotal],
                        'Lifetime Lifetime_refund_rate ':[Lifetime_refund_rate],
                        'Chargebacks for past 180 days':[chargebacks180],
                        '180 day volume':[volume180],
                        '180 day chargeback rate':[day180_chargeback_rate],
                        'Lifetime chargebacks':[chargebackslifetime],
                        'Lifetime volume':[volumetotal],
                        'Lifetime chargeback rate':[Lifetime_chargeback_rate],
                        '90 Days':[days90],
                        '180 Days':[days180],
                        })

        format_mapping = {'Refunds for past 90 days':'${:,.2f}',
                        '90 day volume':'${:,.2f}',
                        '90 day refund rate':'{:.2%}',
                        'Lifetime refunds':'${:,.2f}',
                        'Lifetime volume':'${:,.2f}',
                        'Lifetime Lifetime_refund_rate ':'{:.2%}',
                        'Chargebacks for past 180 days':'${:,.2f}',
                        '180 day volume':'${:,.2f}',
                        '180 day chargeback rate':'{:.2%}',
                        'Lifetime chargebacks':'${:,.2f}',
                        'Lifetime volume':'${:,.2f}',
                        'Lifetime chargeback rate':'{:.2%}',
                        }

        for key, value in format_mapping.items():
            dfcalc[key] = dfcalc[key].apply(value.format)

        df4['total'] = df4['total'].apply('${:,.0f}'.format)

        objects_to_download = {
            "Sheet1": df4,
            "Sheet2": dfcalc,
        }

        components.html(
            download_button(objects_to_download, st.session_state.filename),
            height=0,
        )

if __name__ == "__main__":
    uploaded_files = None
    st.title("Streamlit Example2")

    with st.form("my_form", clear_on_submit=True):
        st.text_input("Filename (must include .xlsx)", key="filename")
        chargebacks180 = st.number_input("Enter Chargebacks For 180 Days", key="chargebacks180")
        chargebackslifetime = st.number_input("Enter Chargebacks for Lifetime", key="chargebackslifetime")
        uploaded_files = st.file_uploader("Upload CSV", type="csv", accept_multiple_files=True)
        submit = st.form_submit_button("Download dataframe")

    if submit:
        download_df()
