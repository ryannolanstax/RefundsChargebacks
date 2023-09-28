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
    try:
        # Create an in-memory Excel writer
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as excel_writer:
            for sheet_name, object_to_download in objects_to_download.items():
                if isinstance(object_to_download, pd.DataFrame):
                    # Write DataFrame as a sheet
                    object_to_download.to_excel(excel_writer, sheet_name=sheet_name)
                else:
                    # Convert other objects to a DataFrame and write as a sheet
                    df = pd.DataFrame({"Data": [object_to_download]})
                    df.to_excel(excel_writer, sheet_name=sheet_name)

        # Seek to the beginning of the in-memory stream
        output.seek(0)
        excel_data = output.read()

        # Encode the Excel file to base64 for download
        b64 = base64.b64encode(excel_data).decode()

        dl_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{download_filename}">Download Excel</a>'

        return dl_link
    except Exception as e:
        # Log the error and return an error message
        st.error(f"An error occurred during file generation: {e}")
        return None

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
        newdf2 = newdf.query("success == 1")
        df3 = newdf2.query("payment_method == 'card' | payment_method == 'bank'")
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

        cardonly = df3.query("type == 'charge' & payment_method == 'card'")
        cardtotal = np.sum(cardonly['total'])
        card180days = np.sum(cardonly[cardonly['created_at'] > days180]['total'])

        Lifetime_refund_rate = (refundtotal / volumetotal)
        day90_refund_rate = (refund90/volume90)
        Lifetime_chargeback_rate = (chargebackslifetime / cardtotal)
        day180_chargeback_rate = (chargebacks180/card180days)


        #####################################

        d = date.today()
        var_names = ["CurrentMonth", "PastMonth", "PastMonth2", "PastMonth3", "PastMonth4", "PastMonth5", "PastMonth6"]
        count = 0
        for name in var_names:
            month, year = (d.month-count, d.year) if d.month != 1 else (12, d.year-1)
            globals()[name] = d.replace(day=1, month=month, year=year)
            count += 1

        volume = df3.query("type == 'charge'")

        volumetotal = np.sum(volume['total'])

        volumeCurrentMonth = np.sum(volume[volume['created_at'] >= CurrentMonth]['total'])

        volumePastMonth = np.sum(volume[volume['created_at'] >= PastMonth]['total']) - volumeCurrentMonth

        volumePastMonth2 = np.sum(volume[volume['created_at'] >= PastMonth2]['total']) - volumeCurrentMonth - volumePastMonth

        volumePastMonth3 = np.sum(volume[volume['created_at'] >= PastMonth3]['total']) - volumeCurrentMonth - volumePastMonth - volumePastMonth2

        volumePastMonth4 = np.sum(volume[volume['created_at'] >= PastMonth4]['total']) - volumeCurrentMonth - volumePastMonth - volumePastMonth2 - volumePastMonth3

        volumePastMonth5 = np.sum(volume[volume['created_at'] >= PastMonth5]['total']) - volumeCurrentMonth - volumePastMonth - volumePastMonth2 - volumePastMonth3 - volumePastMonth4

        volumePastMonth6 = np.sum(volume[volume['created_at'] >= PastMonth6]['total']) - volumeCurrentMonth - volumePastMonth - volumePastMonth2 - volumePastMonth3 - volumePastMonth4 - volumePastMonth5

        volume6monthtotal = np.sum(volume[volume['created_at'] >= PastMonth6]['total'])

        dfsalesvolume = pd.DataFrame({'Past Month 6':[volumePastMonth6],
                                    'Past Month 5':[volumePastMonth5],
                                    'Past Month 4':[volumePastMonth4],
                                    'Past Month 3':[volumePastMonth3],
                                    'Past Month 2':[volumePastMonth2],
                                    'Past Month 1':[volumePastMonth],
                                    'Current Month To Date':[volumeCurrentMonth],
                                    '6 month total':[volume6monthtotal],
                                    'lifetime_total':[volumetotal]
                            }, index=['Sales_Amount'])

        volumecounttotal = np.count_nonzero(volume['total'])

        countCurrentMonth = np.count_nonzero(volume[volume['created_at'] >= CurrentMonth]['total'])

        countPastMonth = np.count_nonzero(volume[volume['created_at'] >= PastMonth]['total']) - countCurrentMonth

        countPastMonth2 = np.count_nonzero(volume[volume['created_at'] >= PastMonth2]['total']) - countCurrentMonth - countPastMonth

        countPastMonth3 = np.count_nonzero(volume[volume['created_at'] >= PastMonth3]['total']) - countCurrentMonth - countPastMonth - countPastMonth2

        countPastMonth4 = np.count_nonzero(volume[volume['created_at'] >= PastMonth4]['total']) - countCurrentMonth - countPastMonth - countPastMonth2 - countPastMonth3

        countPastMonth5 = np.count_nonzero(volume[volume['created_at'] >= PastMonth5]['total']) - countCurrentMonth - countPastMonth - countPastMonth2 - countPastMonth3 - countPastMonth4

        countPastMonth6 = np.count_nonzero(volume[volume['created_at'] >= PastMonth6]['total']) - countCurrentMonth - countPastMonth - countPastMonth2 - countPastMonth3 - countPastMonth4 - countPastMonth5

        counttotalPastMonth6 = np.count_nonzero(volume[volume['created_at'] >= PastMonth6]['total'])

        dfsalescount = pd.DataFrame({'Past Month 6':[countPastMonth6],
                                    'Past Month 5':[countPastMonth5],
                                    'Past Month 4':[countPastMonth4],
                                    'Past Month 3':[countPastMonth3],
                                    'Past Month 2':[countPastMonth2],
                                    'Past Month 1':[countPastMonth],
                                    'Current Month To Date':[countCurrentMonth],
                                    '6 month total':[counttotalPastMonth6],
                                    'lifetime_total':[volumecounttotal]
                            }, index=['Sales_Count'])

        volume90 = np.sum(volume[volume['created_at'] > days90]['total'])

        volume180 = np.sum(volume[volume['created_at'] > days180]['total'])

        #Refund Values

        refund = df3.query("type == 'refund'")

        refundtotal = np.sum(refund['total'])

        refundCurrentMonth = np.sum(refund[refund['created_at'] >= CurrentMonth]['total'])

        refundPastMonth = np.sum(refund[refund['created_at'] >= PastMonth]['total']) - refundCurrentMonth

        refundPastMonth2 = np.sum(refund[refund['created_at'] >= PastMonth2]['total']) - refundCurrentMonth - refundPastMonth

        refundPastMonth3 = np.sum(refund[refund['created_at'] >= PastMonth3]['total']) - refundCurrentMonth - refundPastMonth - refundPastMonth2

        refundPastMonth4 = np.sum(refund[refund['created_at'] >= PastMonth4]['total']) - refundCurrentMonth - refundPastMonth - refundPastMonth2 - refundPastMonth3

        refundPastMonth5 = np.sum(refund[refund['created_at'] >= PastMonth5]['total']) - refundCurrentMonth - refundPastMonth - refundPastMonth2 - refundPastMonth3 - refundPastMonth4

        refundPastMonth6 = np.sum(refund[refund['created_at'] >= PastMonth6]['total']) - refundCurrentMonth - refundPastMonth - refundPastMonth2 - refundPastMonth3 - refundPastMonth4 - refundPastMonth5

        refundtotalPastMonth6 = np.sum(refund[refund['created_at'] >= PastMonth6]['total'])

        dfrefundamount = pd.DataFrame({'Past Month 6':[refundPastMonth6],
                                    'Past Month 5':[refundPastMonth5],
                                    'Past Month 4':[refundPastMonth4],
                                    'Past Month 3':[refundPastMonth3],
                                    'Past Month 2':[refundPastMonth2],
                                    'Past Month 1':[refundPastMonth],
                                    'Current Month To Date':[refundCurrentMonth],
                                    '6 month total':[refundtotalPastMonth6],
                                    'lifetime_total':[refundtotal]
                            }, index=['Refund_Amount'])

        refundcounttotal = np.count_nonzero(refund['total'])

        refundCountCurrentMonth = np.count_nonzero(refund[refund['created_at'] >= CurrentMonth]['total'])

        refundCountPastMonth = np.count_nonzero(refund[refund['created_at'] >= PastMonth]['total']) - refundCountCurrentMonth

        refundCountPastMonth2 = np.count_nonzero(refund[refund['created_at'] >= PastMonth2]['total']) - refundCountCurrentMonth - refundCountPastMonth

        refundCountPastMonth3 = np.count_nonzero(refund[refund['created_at'] >= PastMonth3]['total']) - refundCountCurrentMonth - refundCountPastMonth - refundCountPastMonth2

        refundCountPastMonth4 = np.count_nonzero(refund[refund['created_at'] >= PastMonth4]['total']) - refundCountCurrentMonth - refundCountPastMonth - refundCountPastMonth2 - refundCountPastMonth3

        refundCountPastMonth5 = np.count_nonzero(refund[refund['created_at'] >= PastMonth5]['total']) - refundCountCurrentMonth - refundCountPastMonth - refundCountPastMonth2 - refundCountPastMonth3 - refundCountPastMonth4

        refundCountPastMonth6 = np.count_nonzero(refund[refund['created_at'] >= PastMonth6]['total']) - refundCountCurrentMonth - refundCountPastMonth - refundCountPastMonth2 - refundCountPastMonth3 - refundCountPastMonth4 - refundCountPastMonth5

        refundCounttotalPastMonth6 = np.count_nonzero(refund[refund['created_at'] >= PastMonth6]['total'])

        dfrefundcount = pd.DataFrame({'Past Month 6':[refundCountPastMonth6],
                                    'Past Month 5':[refundCountPastMonth5],
                                    'Past Month 4':[refundCountPastMonth4],
                                    'Past Month 3':[refundCountPastMonth3],
                                    'Past Month 2':[refundCountPastMonth2],
                                    'Past Month 1':[refundCountPastMonth],
                                    'Current Month To Date':[refundCountCurrentMonth],
                                    '6 month total':[refundCounttotalPastMonth6],
                                    'lifetime_total':[refundcounttotal]
                            }, index=['Refund_Count'])

        #Average Ticket

        avgcounttotal = np.average(volume['total'])

        avgCurrentMonth = np.average(volume[volume['created_at'] >= CurrentMonth]['total'])

        avgPastMonth = np.average(volume[(volume['created_at'] >= PastMonth) & (volume['created_at'] < CurrentMonth) ]['total'])

        avgPastMonth2 = np.average(volume[(volume['created_at'] >= PastMonth2) & (volume['created_at'] < PastMonth)  ]['total'])

        avgPastMonth3 = np.average(volume[(volume['created_at'] >= PastMonth3) &  (volume['created_at'] < PastMonth2)]['total'])

        avgPastMonth4 = np.average(volume[(volume['created_at'] >= PastMonth4) &  (volume['created_at'] < PastMonth3)]['total'])

        avgPastMonth5 = np.average(volume[(volume['created_at'] >= PastMonth5) & (volume['created_at'] < PastMonth4)]['total'])

        avgPastMonth6 = np.average(volume[(volume['created_at'] >= PastMonth6) & (volume['created_at'] < PastMonth5)]['total'])

        avgtotalPastMonth6 = np.average(volume[volume['created_at'] >= PastMonth6]['total'])

        dfavgsalescount = pd.DataFrame({'Past Month 6':[avgPastMonth6],
                                    'Past Month 5':[avgPastMonth5],
                                    'Past Month 4':[avgPastMonth4],
                                    'Past Month 3':[avgPastMonth3],
                                    'Past Month 2':[avgPastMonth2],
                                    'Past Month 1':[avgPastMonth],
                                    'Current Month To Date':[avgCurrentMonth],
                                    '6 month total':[avgcounttotal],
                                    'lifetime_total':[avgtotalPastMonth6]
                            }, index=['Average_Ticket'])

        #add in high ticket

        highesttran = np.max(volume['total'])

        highesttranCurrentMonth = np.max(volume[volume['created_at'] >= CurrentMonth]['total'])

        highesttranPastMonth = np.max(volume[(volume['created_at'] >= PastMonth) & (volume['created_at'] < CurrentMonth) ]['total'])

        highesttranPastMonth2 = np.max(volume[(volume['created_at'] >= PastMonth2) & (volume['created_at'] < PastMonth)  ]['total'])

        highesttranPastMonth3 = np.max(volume[(volume['created_at'] >= PastMonth3) &  (volume['created_at'] < PastMonth2)]['total'])

        highesttranPastMonth4 = np.max(volume[(volume['created_at'] >= PastMonth4) &  (volume['created_at'] < PastMonth3)]['total'])

        highesttranPastMonth5 = np.max(volume[(volume['created_at'] >= PastMonth5) & (volume['created_at'] < PastMonth4)]['total'])

        highesttranPastMonth6 = np.max(volume[(volume['created_at'] >= PastMonth6) & (volume['created_at'] < PastMonth5)]['total'])

        highesttrantotalPastMonth6 = np.max(volume[volume['created_at'] >= PastMonth6]['total'])

        dfhighesttrans = pd.DataFrame({'Past Month 6':[highesttranPastMonth6],
                                    'Past Month 5':[highesttranPastMonth5],
                                    'Past Month 4':[highesttranPastMonth4],
                                    'Past Month 3':[highesttranPastMonth3],
                                    'Past Month 2':[highesttranPastMonth2],
                                    'Past Month 1':[highesttranPastMonth],
                                    'Current Month To Date':[highesttranCurrentMonth],
                                    '6 month total':[highesttrantotalPastMonth6],
                                    'lifetime_total':[highesttran]
                            }, index=['Highest_Transactions'])
        
        if volumePastMonth6 > 0: 
            refundamountmonth6 = 100.00 * (refundPastMonth6 / volumePastMonth6)
        else: 
            refundamountmonth6 = 0

        if volumePastMonth5 > 0: 
            refundamountmonth5 = 100.00 * (refundPastMonth5 / volumePastMonth5)
        else: 
            refundamountmonth5 = 0

        if volumePastMonth4 > 0: 
            refundamountmonth4 = 100.00 * (refundPastMonth4 / volumePastMonth4)
        else: 
            refundamountmonth4 = 0

        if volumePastMonth3 > 0: 
            refundamountmonth3 = 100.00 * (refundPastMonth3 / volumePastMonth3)
        else: 
            refundamountmonth3 = 0

        if volumePastMonth2 > 0: 
            refundamountmonth2 = 100.00 * (refundPastMonth2 / volumePastMonth2)
        else: 
            refundamountmonth2 = 0

        if volumePastMonth > 0: 
            refundamountmonth = 100.00 * (refundPastMonth / volumePastMonth)
        else: 
            refundamountmonth = 0

        if volumeCurrentMonth > 0: 
            refundamountcurrentmonth = 100.00 * (refundCurrentMonth / volumeCurrentMonth)
        else: 
            refundamountcurrentmonth = 0

        if volume6monthtotal > 0: 
            refundamount6month = 100.00 * (refundtotalPastMonth6 / volume6monthtotal)
        else: 
            refundamount6month = 0

        if volumetotal > 0: 
            refundamounttotal = 100.00 * (refundtotal / volumetotal)
        else: 
            refundamounttotal = 0

        dfrefundamountpercent = pd.DataFrame({'Past Month 6':[refundamountmonth6],
                             'Past Month 5':[refundamountmonth5],
                             'Past Month 4':[refundamountmonth4],
                             'Past Month 3':[refundamountmonth3],
                             'Past Month 2':[refundamountmonth2],
                             'Past Month 1':[refundamountmonth],
                             'Current Month To Date':[refundamountcurrentmonth],
                             '6 month total':[refundamount6month],
                             'lifetime_total':[refundamounttotal]
                       }, index=['Refund_Amount_Ratio'])
        
        if countPastMonth6 > 0: 
            refundcountpercentmonth6 = 100.00 * (refundCountPastMonth6 / countPastMonth6)
        else: 
            refundcountpercentmonth6 = 0

        if countPastMonth5 > 0: 
            refundcountpercentmonth5 = 100.00 * (refundCountPastMonth5 / countPastMonth5)
        else: 
            refundcountpercentmonth5 = 0

        if countPastMonth4 > 0: 
            refundcountpercentmonth4 = 100.00 * (refundCountPastMonth4 / countPastMonth4)
        else: 
            refundcountpercentmonth4 = 0

        if countPastMonth3 > 0: 
            refundcountpercentmonth3 = 100.00 * (refundCountPastMonth3 / countPastMonth3)
        else: 
            refundcountpercentmonth3 = 0

        if countPastMonth2 > 0: 
            refundcountpercentmonth2 = 100.00 * (refundCountPastMonth2 / countPastMonth2)
        else: 
            refundcountpercentmonth2 = 0

        if countPastMonth > 0: 
            refundcountpercentmonth = 100.00 * (refundCountPastMonth / countPastMonth)
        else: 
            refundcountpercentmonth = 0

        if countCurrentMonth > 0: 
            refundcountpercentcurrentmonth = 100.00 * (refundCountCurrentMonth / countCurrentMonth)
        else: 
            refundcountpercentcurrentmonth = 0

        if volume6monthtotal > 0: 
            refundcountpercent6month = 100.00 * (refundCounttotalPastMonth6 / counttotalPastMonth6)
        else: 
            refundcountpercent6month = 0

        if volumecounttotal > 0: 
            refundcountpercenttotal = 100.00 * (refundcounttotal / volumecounttotal)
        else: 
            refundcountpercenttotal = 0

        dfrefundpercentcount = pd.DataFrame({'Past Month 6':[refundcountpercentmonth6],
                             'Past Month 5':[refundcountpercentmonth5],
                             'Past Month 4':[refundcountpercentmonth4],
                             'Past Month 3':[refundcountpercentmonth3],
                             'Past Month 2':[refundcountpercentmonth2],
                             'Past Month 1':[refundcountpercentmonth],
                             'Current Month To Date':[refundcountpercentcurrentmonth],
                             '6 month total':[refundcountpercent6month],
                             'lifetime_total':[refundcountpercenttotal]
                       }, index=['Refund_Count_Ratio'])
        

        dflastersresults = pd.concat([dfsalesvolume, dfsalescount, dfrefundamount, dfrefundcount, dfavgsalescount, dfhighesttrans,dfrefundamountpercent, dfrefundpercentcount], axis=0)


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
            "Sheet2": dflastersresults,
            "Sheet3": dfcalc,
        }

        download_link = download_button(objects_to_download, st.session_state.filename)
        if download_link:
            st.markdown(download_link, unsafe_allow_html=True)
        else:
            st.error("File download failed.")

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
