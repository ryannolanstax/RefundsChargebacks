import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import xlsxwriter
from datetime import date, timedelta

chargebacks180 = st.number_input("Enter Chargebacks For 180 Days", key="chargebacks180")
chargebackslifetime = st.number_input("Enter Chargebacks for Lifetime", key="chargebackslifetime")

uploaded_files = st.file_uploader("Upload CSV", type="csv", accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file.seek(0)
    uploaded_data_read = [pd.read_csv(file) for file in uploaded_files]
    df = pd.concat(uploaded_data_read)

    buffer = io.BytesIO()


    days90 = date.today() - timedelta(days=90)
    days180 = date.today() - timedelta(days=180)

    df2 = df[df['created_at'].str.split(expand=True)[1].isna() == False]
    dfbaddates = df[df['created_at'].str.split(expand=True)[1].isna() == True]

    dfbaddates['created_at'] = dfbaddates['created_at'].apply(lambda x: datetime.datetime(1900, 1, 1, 0, 0, 0) + datetime.timedelta(days=float(x)))
    dfbaddates['created_at'] = dfbaddates['created_at'].dt.strftime('%m/%d/%y')
    df2['created_at'] = df2['created_at'].str.split(expand=True)[0].str.strip() 
    newdf = pd.concat([df2, dfbaddates])
    newdf['created_at'] = pd.to_datetime(newdf['created_at']).dt.date
    df3 = newdf.query("payment_method == 'card' | payment_method == 'bank'")

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

  #  dfcalc.to_csv('test.csv', index=False)


    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Write each dataframe to a different worksheet.
        df3.to_excel(writer, sheet_name='Clean_Data')
        dfcalc.to_excel(writer, sheet_name='Calculations')

        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()

        st.download_button(
            label="Download Excel worksheets",
            data=buffer,
            file_name="refundchargeback.xlsx",
            mime="application/vnd.ms-excel"
        )

else:
   st.warning("you need to upload a csv file.")
