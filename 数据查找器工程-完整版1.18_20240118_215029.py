#coding=utf-8
import xlwings as xw
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk

app=None

month_list=[]
com_list=[]
field_list=[]
indicator=''
base_path=''
report_path=''
switch_com_dict={'全口径':'集团','集团本部':'母公司','股份本部':'股份','北方':'瓦房店'}
switch_field_dict={'售电量':'上网电量','上网电量':'售电量','耗油量':'耗柴油量','期末供暖收费面积':'面积','阻垢剂':'水质稳定剂','热网补水量':'热网耗水量'}
misunderstanding_dict={'耗水量':'脱硫脱硝耗水量','耗电量':'脱硫脱硝耗电量'}


#0  检查“耗水量”及其它易混指标
def check_mis(field,find_field,misunderstanding_dict):
  for mis_field in misunderstanding_dict:
    if field in mis_field and find_field == misunderstanding_dict[mis_field]:
      return True
  return False

#1-4  日期转换路径程序
def month_to_path(month,base_path):
  if '.0' in month:                       #此处处理误加'.0'的情况，如'23.01'
    month=month.replace('.0','.')
  filename='综合表'+month+'.xls'
  path=os.path.join(base_path,filename)
  return path

#1-5  升级版批量日期转换路径程序
def months_to_paths(month_list,base_path):
  list=[]
  for month in month_list:
    list.append(month_to_path(month,base_path))
  return list

#1-1  最基础的找数据程序。这个程序的定位条件相当宽容，只要含有指定字符串即可，不必完全一致。但使用时也需要注意！！！加入耗水、耗电匹配检测
def to_find_value(wb,com,field,indicator):
  #基础区
  sheet_list=wb.sheets
  sht=None
  value=None
  find_range=None
  #找表
  for sheet in sheet_list:
    if com in sheet.name:
      sht=sheet
      break
  #如果表单存在，就去定位数据及单位所在位置。增加检查因指标包含关系导致找错的情况检查
  if sht != None:
    find_range=sht[:,0:3].api.Find(field)
    #find_range=sht.api.UsedRange.Find(field)
    if find_range != None:
      #检查
      if check_mis(field,str(sht.range(find_range.Row,find_range.Column).value).strip(),misunderstanding_dict) !=False:
        find_range=sht[:,0:3].api.FindNext(find_range)
        #find_range=sht.api.UsedRange.FindNext(find_range)
    
      row_value=find_range.Row
      col_value=sht.api.UsedRange.Find(indicator).Column
      row_unit=row_value
      col_unit=find_range.Column+1
      unit=str(sht.range(row_unit,col_unit).value).strip()
      value=sht.range(row_value,col_value).value
      
      #print(com)
      #print(field)
      #print(value,unit)
      if value is not None:
        return str(value)+str(unit)
      else:
        return '无数据'
    else :
      return '无数据'
  else:
    return '表单不存在'
  
#1-2.1  单位名称处理函数
def com_switch(com,switch_com_dict):
  for key_name in switch_com_dict:
    if com in key_name:
      return switch_com_dict[key_name]
    else:
      continue
  return com

#1-2.2  指标名称处理函数
def field_switch(field,switch_field_dict):
  for field_name in switch_field_dict:
    if field in field_name:
      return switch_field_dict[field_name]
    else:
      continue
  return field

#1-3.1  升级版找数据程序：寻找单位别名
def to_find_value_plus_previous(wb,com,field,indicator,switch_com_dict):
  x=to_find_value(wb,com,field,indicator)
  if x != '表单不存在':
    return x
  else:
    x=to_find_value(wb,com_switch(com,switch_com_dict),field,indicator)
    return x

#1-3.2  升级版找数据程序：寻找指标别名
def to_find_value_plus(wb,com,field,indicator,switch_com_dict,switch_field_dict):
  x=to_find_value_plus_previous(wb,com,field,indicator,switch_com_dict)
  if x=='无数据':
    x=to_find_value_plus_previous(wb,com,field_switch(field,switch_field_dict),indicator,switch_com_dict)
  return x

#1-6  在升级版的基础上，构建批量找数据程序1:单指标&单单位&多路径，结合月份处理程序，内含单位名处理程序
def to_find_more_values(month_list,com,field,indicator,switch_com_dict,switch_field_dict,base_path):
  value_dict={}
  #path_list=months_to_paths(month_list,base_path)
  for month in month_list:
    path=month_to_path(month,base_path)
    wb=app.books.open(path)
    x=to_find_value_plus(wb,com,field,indicator,switch_com_dict,switch_field_dict)
    value_dict[month]=x
    wb.close()
  return value_dict

#1-7  最终找数据程序，接受 多指标&多单位&多月份 ，按单指标，单单位，多月份循环获得结果，内含结合月份处理程序+单位名处理程序
def to_find_all_values(month_list,com_list,field_list,indicator,switch_com_dict,switch_field_dict,base_path):
  value_dict={}
  #path_list=months_to_paths(month_list,base_path)
  for field in field_list:
    value_dict[field]={}      #此处需要先声明value_dict[field]是个字典，否则不能直接在其下级赋值
    for com in com_list:
      value_dict[field][com]=to_find_more_values(month_list,com,field,indicator,switch_com_dict,switch_field_dict,base_path)
      print(field,com,'in')
      print(value_dict[field][com])
  return value_dict


#2-1  输出程序，将程序1-7中生成的结果字典，输出到一个结果文件中
def report(value_dict,path):
  row=1
  col=1
  #app=xw.App(visible=False,add_book=True)
  wb=app.books.add()
  sht=wb.sheets[0]
  #一步一步拆解字典
  for field in value_dict:
    if row>1:
      row=row+2
    sht.range(row,col).value=field
    row=row+1                           #在每次写完内容后，为下次写入做好移动光标的准备
    col=col+1
    #在此时，就应该写上月份
    month_list=list(value_dict[field][list(value_dict[field].keys())[0]].keys())
    for month in month_list:
      sht.range(row,col).value=month
      col=col+1
    row=row+1
    col=1
    for com in value_dict[field]:
      sht.range(row,col).value=com
      col=col+1
      for month in value_dict[field][com]:
        sht.range(row,col).value=value_dict[field][com][month]
        col=col+1
      row=row+1
      col=1
  wb.save(path)
  

# ↓panel2库↓=======================================================================================
base_path2=''
filename_list_text2=''
filename_list2=''
#清理变量file_path2=''
sheet_name2=''
position_list_text2=''
report_path2=''

#0-1 根目录下找文件函数
def to_find_file(root_dir, target_file):
  for root, dirs, files in os.walk(root_dir):
    if target_file in files:
      return os.path.join(root, target_file)
  return None

#0-2 文件列表生成函数
def to_make_file_list(base_path,filename_list_text,file_expand_text):
  filename_list=[]
  if '/' in file_expand_text:
    for file_name in filename_list_text.split('，'):  
      for file_expand in file_expand_text.split('/'):
        if os.path.exists(base_path+'\\'+file_name+file_expand):
          filename_list.append(base_path+'\\'+file_name+file_expand)
  else:

    filename_list=[i+file_expand_text for i in filename_list_text.split('，')]
  return filename_list


#1-1 panel2的基本查找程序
def to_find_value2(file_path2, sheet_name2, position2):
  if file_path2 ==None or not os.path.exists(file_path2):
    return '找不到文件'
  wb = app.books.open(file_path2)
  sheet = wb.sheets[sheet_name2]
  value = sheet.range(position2).value
  wb.close()
  if value!=None:
    return value
  else:
    return '无数据'

#1-2 panel2的进阶批量查找程序
def to_find_more_values2(base_path2,filename_list2,sheet_name2, position2):
  temp_dict={}
  for filename in filename_list2:
    temp_path=to_find_file(base_path2,filename)
    #temp_path=os.path.join(base_path2,filename)
    temp_dict[filename]=to_find_value2(temp_path,sheet_name2, position2)
  return  temp_dict

#1-3 panel2的多位置批量查找程序
def to_find_all_values2(base_path2,filename_list2,sheet_name2, position_list_text2):
  position_list2=position_list_text2.split(',')
  result_dict={}
  for position in position_list2:
    result_dict[position]={}
    result_dict[position]=to_find_more_values2(base_path2,filename_list2,sheet_name2, position)
    print(result_dict)
  return result_dict

#2-1 panel2的输出程序，将程序1-3中生成的结果字典，输出到一个结果文件中
def report2(value_dict2,report_path2):
  wb=app.books.add()
  sht=wb.sheets[0]
  row=2
  for position in value_dict2:
    column=2
    for filename,value in value_dict2[position].items():
      sht.range(row,column).value=filename
      sht.range(row+1,column).value=value
      sht.range(row+1,1).value=position
      column=column+1
    row=row+3
  wb.save(report_path2)

# ----------↑变量及函数区↑--------------------------------------------------------------------------------------------------------------------------------------------------------------------

def f():
  global month_list,com_list,field_list,indicator,base_path,report_path,switch_com_dict,switch_field_dict,app
  #try:
    #变量处理区
  app=xw.App(visible=False,add_book=False)
  button_text_var.set('处理中……')
  month_list_raw=month_list_var.get()
  month_list=month_list_raw.split(',')

  com_list_raw=com_list_var.get()
  com_list=com_list_raw.split('，')

  field_list_raw=field_list_var.get()
  field_list=field_list_raw.split('，')

  indicator=indicator_var.get()
  base_path=base_path_var.get()
  report_path=report_path_var.get()

  print(month_list,com_list,field_list,indicator,base_path,report_path)

  if month_list ==[] or com_list==[] or field_list==[] or indicator=='' or base_path=='' or report_path=='':
    print("输入有误，请重新输入")
    button_text_var.set('输入有误，请重新输入！')
    return None
  else:
    report_path=os.path.join(report_path,'报告.xlsx')
    report(to_find_all_values(month_list,com_list,field_list,indicator,switch_com_dict,switch_field_dict,base_path),report_path)
    button_text_var.set('报告生成成功')
    app.quit()
  #except Exception as e:
  #  button_text_var.set('输入有误，请重新输入！')
  #  print(e.args)
def f2():
  global base_path2,filename_list_text2,file_path2,sheet_name2,position2,position_list_text2,file_expand2,app
  button_text_var2.set('处理中……')
  app=xw.App(visible=False,add_book=False)

  file_expand2=file_expand_var2.get()
  base_path2=base_path_var2.get()
  filename_list_text2=filename_list_var2.get()
  sheet_name2=sheet_name_var2.get()
  #变量清理position2=position_var2.get()
  position_list_text2=position_list_text2_var2.get()
  report_path2=report_path_var2.get()

  filename_list2=to_make_file_list(base_path2,filename_list_text2,file_expand2)
  print(filename_list2)

  report_path2=os.path.join(report_path2,'report.xlsx')
  print('base_path2:'+base_path2)
  #print('filename_list2:'+filename_list2)
  print('sheet_name2:'+sheet_name2)
  print('report_path2:'+report_path2)
  report2(to_find_all_values2(base_path2,filename_list2,sheet_name2,position_list_text2),report_path2)
  app.quit()
  button_text_var2.set('报告生成成功')

# 创建窗口
window = tk.Tk()
window.title("数据提取器")
window.geometry("300x335")

# panel1的联动变量
month_list_var=tk.StringVar()
com_list_var=tk.StringVar()
field_list_var=tk.StringVar()
indicator_var=tk.StringVar()
base_path_var=tk.StringVar()
report_path_var=tk.StringVar()
button_text_var=tk.StringVar()

button_text_var.set('点击生成报告')
month_list_var.set('23.10,23.11,23.12,24.1,24.2,24.3,24.4,24.10,24.11,24.12,25.1,25.2,25.3,25.4')
field_list_var.set('生产耗原煤量，耗标煤总量')
com_list_var.set('全口径，主城区，北海，水炉，香海，金州，北方，金普，庄河')
#field_list_var.set('发电量')
#com_list_var.set('庄河')
indicator_var.set('实际')
base_path_var.set(r'E:\BaiduSyncdisk\工作区\基础数据\采摘数据库（调整）')
#base_path_var.set(r'C:\Users\ww\Desktop')
report_path_var.set(r'C:\Users\ZQB8182\Desktop')
#report_path_var.set(r'C:\Users\ww\Desktop')


#panel2的联动变量
base_path_var2=tk.StringVar()
filename_list_var2=tk.StringVar()
sheet_name_var2=tk.StringVar()
#变量清理position_var2=tk.StringVar()
report_path_var2=tk.StringVar()
position_list_text2_var2=tk.StringVar()
button_text_var2=tk.StringVar()
file_expand_var2=tk.StringVar()

button_text_var2.set('点击使用通用生成器')
base_path_var2.set(r'C:\Users\ZQB8182\Desktop\日报数据库')
report_path_var2.set(r'C:\Users\ZQB8182\Desktop')
filename_list_var2.set('2023.11.5 能源集团生产日报，2023.11.6 能源集团生产日报，2023.11.7 能源集团生产日报，2023.11.8 能源集团生产日报，2023.11.9 能源集团生产日报，2023.11.10 能源集团生产日报，2023.11.11 能源集团生产日报，2023.11.12 能源集团生产日报，2023.11.13 能源集团生产日报，2023.11.14 能源集团生产日报，2023.11.15 能源集团生产日报，2023.11.16 能源集团生产日报，2023.11.17 能源集团生产日报，2023.11.18 能源集团生产日报，2023.11.19 能源集团生产日报，2023.11.20 能源集团生产日报，2023.11.21 能源集团生产日报，2023.11.22 能源集团生产日报，2023.11.23 能源集团生产日报，2023.11.24 能源集团生产日报，2023.11.25 能源集团生产日报，2023.11.26 能源集团生产日报，2023.11.27 能源集团生产日报，2023.11.28 能源集团生产日报，2023.11.29 能源集团生产日报，2023.11.30 能源集团生产日报，2023.12.1能源集团生产日报，2023.12.2 能源集团生产日报，2023.12.3 能源集团生产日报，2023.12.4 能源集团生产日报，2023.12.5 能源集团生产日报，2023.12.6 能源集团生产日报，2023.12.7 能源集团生产日报，2023.12.8 能源集团生产日报，2023.12.9 能源集团生产日报，2023.12.10 能源集团生产日报，2023.12.11 能源集团生产日报，2023.12.12 能源集团生产日报，2023.12.13 能源集团生产日报，2023.12.14 能源集团生产日报，2023.12.15 能源集团生产日报，2023.12.16 能源集团生产日报，2023.12.17 能源集团生产日报，2023.12.18 能源集团生产日报，2023.12.19 能源集团生产日报，2023.12.20 能源集团生产日报，2023.12.21 能源集团生产日报，2023.12.22 能源集团生产日报，2023.12.23 能源集团生产日报，2023.12.24 能源集团生产日报，2023.12.25 能源集团生产日报，2023.12.26 能源集团生产日报，2023.12.27 能源集团生产日报，2023.12.28 能源集团生产日报，2023.12.29 能源集团生产日报，2023.12.30 能源集团生产日报，2023.12.31 能源集团生产日报，2024.1.1 能源集团生产日报，2024.1.2 能源集团生产日报，2024.1.3 能源集团生产日报，2024.1.4 能源集团生产日报，2024.1.5 能源集团生产日报，2024.1.6 能源集团生产日报，2024.1.7 能源集团生产日报，2024.1.8 能源集团生产日报，2024.1.9 能源集团生产日报，2024.1.10 能源集团生产日报，2024.1.11 能源集团生产日报，2024.1.12 能源集团生产日报，2024.1.13 能源集团生产日报，2024.1.14 能源集团生产日报，2024.1.15 能源集团生产日报，2024.1.16 能源集团生产日报，2024.1.17 能源集团生产日报，2024.1.18 能源集团生产日报，2024.1.19 能源集团生产日报，2024.1.20 能源集团生产日报，2024.1.21 能源集团生产日报，2024.1.22 能源集团生产日报，2024.1.23 能源集团生产日报，2024.1.24 能源集团生产日报，2024.1.25 能源集团生产日报，2024.1.26 能源集团生产日报，2024.1.27 能源集团生产日报，2024.1.28 能源集团生产日报，2024.1.29 能源集团生产日报，2024.1.30 能源集团生产日报，2024.1.31 能源集团生产日报')
file_expand_var2.set('.xlsx')
sheet_name_var2.set('分析简报')
#变量清理position_var2.set('')
position_list_text2_var2.set('F21,F22,I21,I22,L21,L22,O21,O22,R21,R22')



# 添加一个notebook及frame
notebook=ttk.Notebook(window)
notebook.pack(expand=True, fill="both")
panel1 = ttk.Frame(notebook)
panel2 = ttk.Frame(notebook)
notebook.add(panel1, text="综合表工具")
notebook.add(panel2, text="通用工具")

# 在panel1上添加6个输入框和对应的标签
label101 = tk.Label(panel1, text='统计指标 (全角分隔)')
label101.pack(fill='x',padx=20)
entry101 = tk.Entry(panel1,textvariable=field_list_var)
entry101.pack(fill='x',padx=20)

label102 = tk.Label(panel1, text='统计单位 (全角分隔)')
label102.pack(fill='x',padx=20)
entry102 = tk.Entry(panel1,textvariable=com_list_var)
entry102.pack(fill='x',padx=20)

label103 = tk.Label(panel1, text='统计月份/年度')
label103.pack(fill='x',padx=20)
entry103 = tk.Entry(panel1,textvariable=month_list_var)
entry103.pack(fill='x',padx=20)

label104 = tk.Label(panel1, text='统计时期')
label104.pack(fill='x',padx=20)
entry104 = tk.Entry(panel1,textvariable=indicator_var)
entry104.pack(fill='x',padx=20)

label105 = tk.Label(panel1, text='数据路径')
label105.pack(fill='x',padx=20)
entry105 = tk.Entry(panel1,textvariable=base_path_var)
entry105.pack(fill='x',padx=20)

label106 = tk.Label(panel1, text='输出路径')
label106.pack(fill='x',padx=20)
entry106 = tk.Entry(panel1,textvariable=report_path_var)
entry106.pack(fill='x',padx=20)
# 添加按钮
button101 = tk.Button(panel1, command=f,textvariable=button_text_var)
button101.pack(fill='x',padx=60,pady=3)

# 在panel2上添加组件
label201 = tk.Label(panel2, text='基本路径')
label201.pack(fill='x',padx=20)
entry201 = tk.Entry(panel2,textvariable=base_path_var2)
entry201.pack(fill='x',padx=20)

label202 = tk.Label(panel2, text='文件名列表 (全角分隔)')
label202.pack(fill='x',padx=20)
entry202 = tk.Entry(panel2,textvariable=filename_list_var2)
entry202.pack(fill='x',padx=20)

label203 = tk.Label(panel2, text='文件名扩展名 (用/分隔)')
label203.pack(fill='x',padx=20)
entry203 = tk.Entry(panel2,textvariable=file_expand_var2)
entry203.pack(fill='x',padx=20)

label204 = tk.Label(panel2, text='表单名称')
label204.pack(fill='x',padx=20)
entry204 = tk.Entry(panel2,textvariable=sheet_name_var2)
entry204.pack(fill='x',padx=20)

label205 = tk.Label(panel2, text='单元格位置')
label205.pack(fill='x',padx=20)
entry205 = tk.Entry(panel2,textvariable=position_list_text2_var2)
entry205.pack(fill='x',padx=20)

label206 = tk.Label(panel2, text='输出路径')
label206.pack(fill='x',padx=20)
entry206 = tk.Entry(panel2,textvariable=report_path_var2)
entry206.pack(fill='x',padx=20)

button201 = tk.Button(panel2, command=f2,textvariable=button_text_var2)
button201.pack(fill='x',padx=60,pady=3)



# 运行窗口
window.mainloop()



'''
def find_next_text(range_to_search, text, start_cell=None):
    """
    在指定的Excel区域中查找文本。

    :param range_to_search: 要搜索的xlwings Range对象。
    :param text: 要查找的文本。
    :param start_cell: 开始查找的单元格，默认为None，即从range_to_search的开始处查找。
    :return: 找到的包含文本的xlwings Range对象，如果未找到则返回None。
    """
    # 如果提供了开始单元格，则从该单元格开始搜索
    if start_cell:
        current_cell = start_cell
    else:
        current_cell = range_to_search.cells(1, 1)

    for i in range(current_cell.row, range_to_search.last_cell.row + 1):
        for j in range(current_cell.column, range_to_search.last_cell.column + 1):
            cell = range_to_search.sheet.range(i, j)
            if cell.value == text and (i > current_cell.row or j > current_cell.column):
                return cell
    return None

'''
