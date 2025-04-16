import os
import pandas as pd
from pathlib import Path
from win32com import client
from tkinter import messagebox
import tkinter as tk
from tqdm import tqdm

def dataPreparation(file_dir):
    data=pd.read_excel(file_dir+'\\Cover.xlsx',dtype='str')

    data['Booking Number']=data['Booking Number'].apply(lambda x:x.replace(' ',''))
    data['Order Number']=data['Order Number'].apply(lambda x:x.replace(' ',''))

    data['bookNoIsMul'] = data['Booking Number'].apply(lambda x: len(x.split(',')))
    mulBookNoDf = data[data['bookNoIsMul'] == 2]
    data = data[data['bookNoIsMul'] == 1]
    data['Folder Name']=data['Booking Number']

    mulBookNoDf['Booking Number1'] = mulBookNoDf['Booking Number'].apply(lambda x: x.split(',')[0])
    mulBookNoDf['Booking Number2'] = mulBookNoDf['Booking Number'].apply(lambda x: x.split(',')[1])

    mulBookNoDf['Folder Name']=mulBookNoDf.apply(lambda x:x['Booking Number1'].split('-')[1] if x['Order Number'][-2:]=='00' else x['Booking Number2'].split('-')[1],axis=1)
    del mulBookNoDf['Booking Number2'], mulBookNoDf['Booking Number1']

    data = pd.concat([data, mulBookNoDf])
    del data['bookNoIsMul'],data['Booking Number']
    data = data.sort_values('Order Number').reset_index(drop=True)
    data=data.set_index('Order Number')
    
    return data


def getNewFileName(data_df):
    folder_name_dict=data_df.to_dict()['Folder Name']
    return folder_name_dict


def renameAndConvert(file_dir, folder_name_dict):
    file_list = Path(file_dir+'\\CCD\\').glob('*[0-9].xlsx')  

    # 打开excel程序
    excel = client.DispatchEx('Excel.Application')

    for i in file_list:
        print()
        print('Converting '+i.name+':')

        new_file_dir=file_dir+'\\'+folder_name_dict[i.name[:-5]]
        os.makedirs(new_file_dir)

        file = excel.Workbooks.Open(i.with_name(i.name))
        flag=True
        for _sheet in tqdm(file.Sheets,desc='Progress'):
            if flag:
                new_file_name='CI'
                flag=False
            else:
                new_file_name='PL'
                flag=True
            _sheet.ExportAsFixedFormat(0, f'{new_file_dir}\\{new_file_name}')
            
        file.Close()

    excel.Quit()


if __name__ == '__main__':
    root = tk.Tk()
    root.wm_attributes('-topmost', 1)
    root.withdraw()
    messagebox.showinfo("提示", "Starting...")

    file_dir = os.getcwd()
    data = dataPreparation(file_dir=file_dir)
    folder_name_dict = getNewFileName(data_df=data)
    renameAndConvert(file_dir=file_dir, folder_name_dict=folder_name_dict)
    
    messagebox.showinfo("提示", "Completed.")
