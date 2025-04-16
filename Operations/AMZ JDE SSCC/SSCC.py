import pandas as pd
from tkinter import messagebox
import tkinter as tk
import os
import math
from openpyxl.styles import Font, Alignment, colors, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl import load_workbook
from datetime import datetime, timedelta
import warnings
from pandas.errors import SettingWithCopyWarning

#
# update log:
# 1. 2025-03-18: 由于sscc数据中的Pack SSCC不一定前两位为00，所以在读取数据时修改补全功能
# 
def data_extraction(file_dir, sscc_data_file_name,sa_data_file_name):
    # ------------------------------------------------------------sscc------------------------------------------------------------
    sscc_data=pd.read_excel(file_dir+sscc_data_file_name,dtype={'Pack SSCC':str})
    sscc_data['Pack SSCC']=sscc_data['Pack SSCC'].apply(lambda x:'00'+x if len(x)==18 else '00000'+x)
    #sscc_data['Pack SSCC']='00'+sscc_data['Pack SSCC']
    print('SSCC data reading complete.')

    # ------------------------------------------------------------sa------------------------------------------------------------
    sa_data=pd.read_excel(file_dir+sa_data_file_name,)
    sa_unfold_data=pd.DataFrame()
    for index,row in sa_data.iterrows():
        times=row['To']-row['From']+1
        start=row['From']
        for i in range(times):
            row['Label Number']=start
            row['Cartons']=times
            sa_unfold_data=pd.concat([sa_unfold_data,pd.DataFrame(row).T],axis=0)
            start+=1
    print('SA data extracting complete.')

    return sscc_data,sa_unfold_data


def calculation(sscc_data,sa_data):
    cols_in_sscc=['Customer PO','Customer/Supplier Item Number', 'Pack SSCC', 'Units Per Container','Label Number',]
    del sa_data['From'],sa_data['To']
    # cols_in_sa=['Booking_Key', 'PO_No', 'ASIN', 'Container_No','Label Number', 'Cartons','serial_no','Lot_Number','Expiry_Date']

    res=pd.merge(sa_data,sscc_data[cols_in_sscc],left_on=['PO_No','ASIN','Label Number'],right_on=['Customer PO','Customer/Supplier Item Number','Label Number'])
    res=res[['Booking_Key', 'Container_No','PO_No', 'ASIN', 'Units Per Container', 'Cartons','Pack SSCC','serial_no','Lot_Number','Expiry_Date']]
    res.rename(columns={
        'Units Per Container':'Units',
        'Pack SSCC':'SSCC'
    },inplace=True)

    return res


def write_excel(config, result, file_dir):
    for key in config.keys():
        print('------------------------------------------')
        print('Writing '+key+' Part:')

        data=result[config[key]['condition']]
        col=config[key]['col']
        for value in data[col].unique():
            file_name=value
            df=data[data[col]==value]

            # --------content--------
            file_path=file_dir+file_name+'.xlsx'
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            df.to_excel(writer,index=False)

            # --------format--------
            worksheet = writer.sheets['Sheet1']

            # 所有行（表格内容）字体设置
            for i in range(1, len(df)+1):
                for cell in worksheet[i]:
                    cell.font=Font(name="Arail", size=11)
            # 列宽设置
            for col_index in range(ord('A'),ord('A')+10+1):
                worksheet.column_dimensions[chr(col_index)].width=21.5

            writer.close()
            print(file_name,' writing complete.')

if __name__ == "__main__":
    root = tk.Tk()
    root.wm_attributes("-topmost", 1)
    root.withdraw()
    messagebox.showinfo("提示", "Starting...")

    folder_path = os.getcwd()
    # folder_path = os.getcwd()+'\\AMZ JDE SSCC'
    # print(folder_path)
    sscc_data_file_name='SSCC.xlsx'
    sa_data_file_name='SA.xlsx'

    sscc_data, sa_data = data_extraction(
        file_dir=folder_path, sscc_data_file_name="\\" + sscc_data_file_name,sa_data_file_name="\\" + sa_data_file_name
    )
    result=calculation(sscc_data,sa_data)
    res_dir = folder_path + "\\RESULT\\"
    config={
    'CY':
    {
        'condition':~result['serial_no'].isna(),
        'col':'Container_No',
    },
    'CFS':
    {
        'condition':result['serial_no'].isna(),
        'col':'Booking_Key',
    },
}
    write_excel(
        config,
        result,
        file_dir=res_dir,
    )
    messagebox.showinfo("提示", "Completed.")