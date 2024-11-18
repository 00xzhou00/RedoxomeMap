import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import numpy as np
from PIL import Image,ImageTk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
inmzfilename=''
outfilename=''
runmode=''
HMDB5csvfile="HMDB5.csv"
#获取输入文件
def getuploadfilename():
    infile=filedialog.askopenfilename(filetypes=(('Excel files','*.xlsx'),))
    inpath_text.set(infile)
    global inmzfilename
    inmzfilename=infile
#获取输出文件
def getoutfilename():
    outfile=filedialog.askopenfilename(filetypes=(('Excel files','*.xlsx'),))
    outpath_text.set(outfile)
    global outfilename
    outfilename=outfile

#匹配m/z
def run():
    infilename=inmzfilename
    min = float(entry_min.get())
    max = float(entry_max.get())
    error = float(entry_error.get())
    reoutybmz = []
    reoutmode=[]
    reouterror = []
    reoutaccession = []
    reoutmonisotopic_molecular_weight = []
    reoutiupac_name = []
    reoutname = []
    reoutchemical_formula = []
    reoutkegg = []

    proutybmz = []
    proutmode = []
    prouterror = []
    proutaccession = []
    proutmonisotopic_molecular_weight = []
    proutiupac_name = []
    proutname = []
    proutchemical_formula = []
    proutkegg = []

    hmdbdata = pd.read_csv(HMDB5csvfile,usecols=['accession', 'monisotopic_molecular_weight', 'iupac_name', 'name','chemical_formula', 'kegg'])
    inmzdata = pd.read_excel(infilename, usecols=['m/z'])
    inmz = inmzdata['m/z'].values.tolist()
    hmdbmz = hmdbdata['monisotopic_molecular_weight'].values.tolist()
#进度条
    pro=tk.ttk.Progressbar(windows)
    pro.place(x=270,y=450,width=300,height=16)
    pro['maximum']=len(hmdbmz)*2
    pro['value']=0
#按参数匹配
    for i in range(len(hmdbmz)) :
        if hmdbmz[i]<min or hmdbmz[i]>max:
            pro['value']=i
            windows.update()
            continue
        else:
            for k in range(len(inmz)):
                if hmdbmz[i]-error<inmz[k]<hmdbmz[i]+error:
                    wcc=abs(hmdbmz[i]-inmz[k])

                    reoutybmz.append(inmz[k])
                    reoutmode.append("Redox")
                    reouterror.append(wcc)
                    reoutaccession.append(hmdbdata.iloc[i,0])
                    reoutmonisotopic_molecular_weight.append(hmdbmz[i])
                    reoutiupac_name.append(hmdbdata.iloc[i,2])
                    reoutname.append(hmdbdata.iloc[i,3])
                    reoutchemical_formula.append(hmdbdata.iloc[i,4])
                    reoutkegg.append(hmdbdata.iloc[i,5])

    reoutdf = pd.DataFrame({'mz': reoutybmz, 'Monisotopic_Mass': reoutmonisotopic_molecular_weight,'Error(Da)': reouterror, 'Mode':reoutmode,'IUPAC_Name': reoutiupac_name, 'Name': reoutname, 'Accession Number': reoutaccession,'Chemical_Formula': reoutchemical_formula, 'KEGG': reoutkegg})

    for i in range(len(hmdbmz)):
        if hmdbmz[i] < min or hmdbmz[i] > max:
            pro['value'] = i+len(hmdbmz)
            windows.update()
            continue
        else:
            for k in range(len(inmz)):
                if hmdbmz[i] - error < inmz[k] + 1.0078 < hmdbmz[i] + error:
                    wcc = abs(hmdbmz[i] - inmz[k] - 1.0078)

                    proutybmz.append(inmz[k])
                    proutmode.append("Protonation")
                    prouterror.append(wcc)
                    proutaccession.append(hmdbdata.iloc[i, 0])
                    proutmonisotopic_molecular_weight.append(hmdbmz[i])
                    proutiupac_name.append(hmdbdata.iloc[i, 2])
                    proutname.append(hmdbdata.iloc[i, 3])
                    proutchemical_formula.append(hmdbdata.iloc[i, 4])
                    proutkegg.append(hmdbdata.iloc[i, 5])
    proutdf = pd.DataFrame({'mz': proutybmz, 'Monisotopic_Mass': proutmonisotopic_molecular_weight, 'Error(Da)': prouterror,'Mode':proutmode,'IUPAC_Name': proutiupac_name, 'Name': proutname, 'Accession Number': proutaccession,'Chemical_Formula': proutchemical_formula, 'KEGG': proutkegg})
    pro['value'] = 0
    pro.destroy()

    acc1 = proutdf['Accession Number'].tolist()
    acc2 = reoutdf['Accession Number'].tolist()
    flag2 = []
    find = 0
    for i in range(len(acc2)):
        for j in range(len(acc1)):
            if acc2[i] == acc1[j]:
                find = 1
                break
        if find == 0:
            flag2.append(0)
            find = 0
        elif find == 1:
            flag2.append(1)
            find = 0

    reoutdf['find'] = flag2
    df2_dropped = reoutdf.drop(reoutdf[reoutdf['find'] == 1].index)
    df2_dropped = df2_dropped.drop(columns='find')
    ou = pd.concat([proutdf, df2_dropped], axis=0)
    ou.sort_values(by="mz", inplace=True, ascending=True)
    merge_cells(ou,['mz','Monisotopic_Mass'],outfilename)


    tk.messagebox.showinfo(title='',message='Completed successfully')

def merge_cells(df, key, output_path=None):
    """
    key 列去重并合并单元格并居中
    Args:
        df: DataFrame输入表
        key: （多个）列名
        output_path: 保存路径

    Returns: Workbook 工作簿

    """
    wb = Workbook()  # 创建工作簿
    ws = wb.active  # 获取第一个工作表

    # 把 key 列 调整到最前面，并进行排序
    col = key if isinstance(key, list) else [key]
    set_col = set(col)
    columns = [*col, *(i for i in df.columns if i not in set_col)]
    _df = df[columns]
    _df.sort_values(key, inplace=True)

    # 将每行数据写入工作表中
    for row in dataframe_to_rows(_df, index=False, header=True):
        ws.append(row)

    align = Alignment(horizontal="center", vertical="center")  # 居中样式
    idx = {-1, _df.shape[0] - 1}
    for i, _ in enumerate(col):
        c = _df[_].values
        idx.update(np.where(c[1:] != c[:-1])[0])
        sorted_idx = sorted(idx)
        for start, end in zip(sorted_idx[:-1], sorted_idx[1:]):

            ws.merge_cells(start_row=start + 3, end_row=end + 2, start_column=i + 1, end_column=i + 1)

            ws.cell(start + 3, i + 1).alignment = align

    if output_path:

        wb.save(output_path)

    return wb



#建立窗口
windows=tk.Tk()
windows.title('')
windows.geometry('800x500')

photo=Image.open("background1.png")
photo=photo.resize((800,500))
backgroundphoto=ImageTk.PhotoImage(photo)
background_label=tk.Label(windows,image=backgroundphoto)
background_label.pack()

text_uploadcsvfile=tk.Label(windows,text='MZ data',font=('Arial',16))
text_uploadcsvfile.place(x=50,y=140)

uploadbutton=tk.Button(windows,text='Select file',command=getuploadfilename,width=10,height=1)
uploadbutton.pack()
uploadbutton.place(x=500,y=140)

inpath_text=tk.StringVar()
entry_inpath=tk.Entry(windows,textvariable=inpath_text,width=45)
entry_inpath.place(x=155,y=145)

text_Da=tk.Label(windows,text='(Da)')
text_Da.place(x=350,y=200)
text_Da2=tk.Label(windows,text='(Da)')
text_Da2.place(x=650,y=200)
text_min=tk.Label(windows,text='Mass range',font=('Arial',16))
text_min.place(x=50,y=200)
text_max=tk.Label(windows,text='to',font=('Arial',16))
text_max.place(x=260,y=200)
text_error=tk.Label(windows,text='Error',font=('Arial',16))
text_error.place(x=500,y=200)
entry_min=tk.Entry(windows,width=5)
entry_min.pack()
entry_min.place(x=200,y=205)
entry_max=tk.Entry(windows,width=5)
entry_max.pack()
entry_max.place(x=300,y=205)
entry_error=tk.Entry(windows,width=8)
entry_error.pack()
entry_error.place(x=570,y=205)
text_outfile=tk.Label(windows,text='Results',font=('Arial',16))
text_outfile.place(x=50,y=260)
outpath_text=tk.StringVar()
surebutton=tk.Button(windows,text='Save to',command=getoutfilename,width=10,height=1)
surebutton.pack()
surebutton.place(x=650,y=260)
entry4=tk.Entry(windows,textvariable=outpath_text,width=65)
entry4.place(x=150,y=265)


runbutton=tk.Button(windows,text='Run',command=run,width=10,height=1)
runbutton.pack()
runbutton.place(x=380,y=400)

windows.mainloop()