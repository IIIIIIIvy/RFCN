import os
import pandas as pd
from pathlib import Path
from win32com import client
from tkinter import messagebox
import tkinter as tk

def read_jde_data(jde_file_name):
    jde_data=pd.read_excel(jde_file_name,usecols="A:D",dtype={'Document Number':str,'Sales Order Number':str})
    jde_data.fillna('',inplace=True)
    jde_data=jde_data.drop(len(jde_data)-1)
    
    new_jde_data=jde_data.groupby('Reference')['Document Number'].apply(lambda x:"/".join(list(x))).reset_index()
    new_jde_data=pd.merge(new_jde_data,jde_data.groupby('Reference')['Sales Order Number'].apply(lambda x:"/".join(list(x))).reset_index())
    new_jde_data=pd.merge(new_jde_data,jde_data.groupby('Reference')['Open Amount'].sum().reset_index())
    return new_jde_data

   

if __name__ == '__main__':
    root = tk.Tk()
    root.wm_attributes('-topmost', 1)
    root.withdraw()
    messagebox.showinfo("提示", "程序启动中！")

    new_folder_dir = os.getcwd()+'\\result\\'
    jde_file_name = os.getcwd()+"\\JDE.xlsx"
    jde_data=read_jde_data(jde_file_name)

    file_list = Path(os.getcwd()+'\\raw\\').glob('*账务明细.csv')
    for i in file_list:
        print(i)
        # ------读取数据部分
        header_df=pd.read_csv(i,sep=",", encoding='gbk',nrows=4,names=['header'])
        content_df=pd.read_csv(i,sep=",", encoding='gbk',header=4)

        # ------计算部分
        content_df_copy=content_df[:-4].copy()
        content_df_copy=content_df_copy[['发生时间','业务基础订单号','收入金额（+元）',
           '支出金额（-元）', '账户余额（元）', '交易渠道', '业务类型', '备注','业务描述' ]]
        for col in content_df_copy.columns:
            if type(content_df_copy.loc[0,col])!=str:
                continue
            content_df_copy[col]=content_df_copy[col].apply(lambda x:x.replace('\t',''))

        df1=content_df_copy[content_df_copy['业务描述'].apply(lambda x:x.find('0010001')!=-1)]
        df2=content_df_copy[content_df_copy['业务描述'].apply(lambda x:x.find('0010001')==-1)]
        df1=pd.merge(df1,jde_data,left_on='业务基础订单号',right_on='Reference',how='left')
        del df1['Reference']
        res1=pd.concat([df1,df2],axis=0)
        res1['Diff']=res1['收入金额（+元）']-res1['Open Amount']

        income_df=res1.groupby('业务描述')['收入金额（+元）'].sum().reset_index(name='Sum of 收入金额')
        payment_df=res1.groupby('业务描述')['支出金额（-元）'].sum().reset_index(name='Sum of 支出金额')
        open_df=res1.groupby('业务描述')['Open Amount'].sum().reset_index(name='Sum of Open Amount')
        diff_df=res1.groupby('业务描述')['Diff'].sum().reset_index(name='Sum of Diff')
    
        res2=pd.merge(income_df,payment_df)
        res2=pd.merge(res2,open_df)
        res2=pd.merge(res2,diff_df)

        new_file_path=new_folder_dir+i.name.replace('csv','xlsx')
        writer= pd.ExcelWriter(new_file_path)
    
        header_df.to_excel(excel_writer=writer,sheet_name='账务明细',index=False, header=False)
        res1.to_excel(excel_writer=writer,sheet_name='diff明细',index=False)
        res2.to_excel(excel_writer=writer,sheet_name='数据统计表',index=False)
        writer.close()

        with pd.ExcelWriter(new_file_path, engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
    	    content_df.to_excel(writer,sheet_name='账务明细',startrow=4, index=False)
    
    messagebox.showinfo("提示",
                        "完成！")