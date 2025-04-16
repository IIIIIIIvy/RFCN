import os
import pandas as pd
from pathlib import Path
from win32com import client
from tkinter import messagebox
import tkinter as tk

def dataPreparation(file_dir):
    data_file_name = '\\JDE.xlsx'
    data_file_name = file_dir + data_file_name
    print(data_file_name)

    df1 = pd.read_excel(data_file_name, sheet_name='Sheet1')
    df1.fillna('', inplace=True)
    df2 = pd.read_excel(data_file_name, sheet_name='Sheet2')
    df2.fillna('', inplace=True)

    data = pd.merge(df1, df2)[['Supplier Name', 'Port', 'Booking Number', 'Customer PO', 'Order Number']]
    data['Booking Number'] = data['Booking Number'].astype(str).apply(lambda x: x.replace(' ', ''))

    data['Remark label'] = data['Booking Number'].apply(lambda x: x.find('BK'))
    data = data[data['Remark label'] == -1]
    del data['Remark label']

    data['bookNoIsMul'] = data['Booking Number'].apply(lambda x: len(x.split(',')))
    mulBookNoDf = data[data['bookNoIsMul'] == 2]
    data = data[data['bookNoIsMul'] == 1]
    mulBookNoDf['Booking Number1'] = mulBookNoDf['Booking Number'].apply(lambda x: x.split(',')[0])
    mulBookNoDf['Booking Number2'] = mulBookNoDf['Booking Number'].apply(lambda x: x.split(',')[1])
    mulBookNoDf = pd.concat([mulBookNoDf[['Supplier Name', 'Port', 'Booking Number', 'Customer PO',
                                          'Order Number', 'bookNoIsMul', 'Booking Number1']],
                             mulBookNoDf[['Supplier Name', 'Port', 'Booking Number', 'Customer PO',
                                          'Order Number', 'bookNoIsMul', 'Booking Number2']]])
    mulBookNoDf.loc[mulBookNoDf['Booking Number2'].isna(), 'Booking Number'] = mulBookNoDf.loc[
        mulBookNoDf['Booking Number2'].isna(), 'Booking Number1']
    mulBookNoDf.loc[mulBookNoDf['Booking Number2'].isna(), 'new Order Number'] = mulBookNoDf.loc[
        mulBookNoDf['Booking Number2'].isna(), 'Order Number'].apply(lambda x: str(x) + '-00')

    mulBookNoDf.loc[mulBookNoDf['Booking Number1'].isna(), 'Booking Number'] = mulBookNoDf.loc[
        mulBookNoDf['Booking Number1'].isna(), 'Booking Number2']
    mulBookNoDf.loc[mulBookNoDf['Booking Number1'].isna(), 'new Order Number'] = mulBookNoDf.loc[
        mulBookNoDf['Booking Number1'].isna(), 'Order Number'].apply(lambda x: str(x) + '-01')
    mulBookNoDf['Order Number'] = mulBookNoDf['new Order Number']
    del mulBookNoDf['new Order Number'], mulBookNoDf['Booking Number2'], mulBookNoDf['Booking Number1']

    data = pd.concat([data, mulBookNoDf])
    data['Order Number'] = data['Order Number'].astype(str)
    data = data.sort_values('Order Number').reset_index(drop=True)
    return data


def getNewFileName(file_dir, data_df):
    folder_path = Path(file_dir)
    file_list = folder_path.glob('*[0-9].xlsx')  # 获取该文件夹下主名以“月销售表”结尾的所有工作簿

    file_name_list = []
    for i in file_list:
        file_name_list.append(i.name[:-5])

    file_name_df = pd.DataFrame(file_name_list, columns=['Order Number'])
    file_name_df = pd.merge(file_name_df, data_df)
    file_name_df['file_name'] = file_name_df['Booking Number'].apply(lambda x: x[-6:]) + ';' + file_name_df[
        'Customer PO'].apply(lambda x: x[:-5]) + ';' + file_name_df['Port']
    file_name_df['file_name'].astype(str)
    return file_name_df


def renameAndConvert(file_dir, file_name_df):
    folder_path = Path(file_dir)
    file_list = folder_path.glob('*[0-9].xlsx')  # 获取该文件夹下主名以“月销售表”结尾的所有工作簿

    # 打开excel程序
    excel = client.DispatchEx('Excel.Application')
    fail_count = 0
    success_count = 0
    fail_file_name = []
    for i in file_list:
        if file_name_df[file_name_df['Order Number'] == i.name[:-5]].empty:
            fail_count = fail_count + 1
            fail_file_name.append(i.name)
            continue
        new_name = file_name_df.loc[file_name_df['Order Number'] == i.name[:-5], 'file_name'].iloc[0]

        old_file_name = i.name
        new_file_name = old_file_name.replace(i.name[:-5], new_name)  # 替换旧的文件名，变成以“月”为结尾的工作簿
        new_file_name = new_file_name.replace('.xlsx', '.pdf')

        old_file_path = i.with_name(old_file_name)
        file = excel.Workbooks.Open(old_file_path)

        # word文件另存为当前文件夹下的pdt文件，0为pdf，1为XPS
        file.ExportAsFixedFormat(0, f'{folder_path}\\{new_file_name}')
        # 关闭word文件
        file.Close()
        success_count = success_count + 1

    # 关闭word程序
    excel.Quit()

    messagebox.showinfo("提示",
                        "完成对" + str(success_count + fail_count) + "个Excel文件的重命名与PDF导出操作，其中成功" + str(
                            success_count) + "个，失败" + str(fail_count) + "个：" + ",".join(fail_file_name))


if __name__ == '__main__':
    root = tk.Tk()
    root.wm_attributes('-topmost', 1)
    root.withdraw()
    messagebox.showinfo("提示", "程序启动中！")

    file_dir = os.getcwd()
    data = dataPreparation(file_dir=file_dir)
    file_name_df = getNewFileName(file_dir=file_dir, data_df=data)
    renameAndConvert(file_dir=file_dir, file_name_df=file_name_df)
