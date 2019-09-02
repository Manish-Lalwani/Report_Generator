#imports

import xlwt
import xlrd
import tkinter
from matplotlib import pyplot as pl
import os
#import tktable
from tkinter import ttk
import pdb
import array


#from tkintertable.Tables import TableCanvas
#from tkintertable.TableModels import TableModel
#from tkintertable import TableCanvas, TableModel
#import tkintertable


def file_names():
	return var_sheet_name_array

def open_workbook(path):
	try:
		book =xlrd.open_workbook(path)
		print("Workbook open return code 0")
		no_of_sheets = book.nsheets
		print ("No of sheets",no_of_sheets)
		return book
	except :
		print("Workbook open return code 1")

def path():
	
    path = "excel_path..."
           

    if (os.path.exists(path)) == True:
	    print("Path return code 0")
	    return path
    else:
        print("Path return code 1")
        return 1
	

def fetch_sheet_by_index(index,book):
	try:
		sheet = book.sheet_by_index(index)
		return sheet
	except:
		print("fetch_sheet return code 1")

def fetch_sheet_by_name(name,book):
	try:
		#sheet = book.fetch_sheet_by_name(name)
		sheet = book.sheet_by_name(name)
		return sheet
	except:
		print("error in fetching sheet fetch sheet return code 1")

def passing_sheet_name_to_fetch_sheet_by_name(book):
	#try:
	sheet_vInfo = fetch_sheet_by_name("vInfo",book)
	sheet_vtools = fetch_sheet_by_name("vTools",book)
	print("success passing")

	array_sheet=[sheet_vInfo,sheet_vtools]
	print("\n\n","Test : func(passing_sheet_name_to_fetch_sheet_by_name) : type of sheet_array",type(sheet_array),"\n\n",)
	#return sheet_vInfo,sheet_vtools #array order
	print("\n\n\n","Printing array sheet data",array_sheet[0])
	return array_sheet
	#except:
	#print("fetch sheet return code 1 //passing")

def reading_column_from_spreadsheet(sheet_array,column_index,sheet_array_index):
	column_val = sheet_array[sheet_array_index].col_values(column_index)

	return column_val
	#i = 0
	#for rows in sheet_array[sheet_array_index]:
		#sheet_array[sheet_array_index].row_values(i,column_index)
		#i++


def filtering_required_columns_from_fetched_sheets(sheet_array):
	VInfo_Vm_column = reading_column_from_spreadsheet(sheet_array, 0, 0)
	print("\n\n","Printing VInfo_Vm_column : \n",VInfo_Vm_column, ",")
	VInfo_HeartBeat_column = reading_column_from_spreadsheet(sheet_array, 7, 0)
	print("\n\n","Printing VInfo_HeartBeat_column : \n",VInfo_HeartBeat_column, ",")
	VInfo_HWVersion_column = reading_column_from_spreadsheet(sheet_array, 47, 0)
	print("\n\n","Printing VInfo_HWVersion_column : \n",VInfo_HWVersion_column, ",")
	VInfo_OSaccordingtotheVMwareTools_column = reading_column_from_spreadsheet(sheet_array, 66, 0)
	print("\n\n","Printing VInfo_OSaccordingtotheVMwareTools_column : \n",VInfo_OSaccordingtotheVMwareTools_column, ",")
	VTools_Tools_column = reading_column_from_spreadsheet(sheet_array, 5, 1)
	print("\n\n","Printing VTools_Tools_column : \n",VTools_Tools_column, ",")

	filtered_column_array=[VInfo_Vm_column,VInfo_HeartBeat_column,VInfo_HWVersion_column,VInfo_OSaccordingtotheVMwareTools_column,VTools_Tools_column]
	return filtered_column_array


def comparing_columns_data():
	print("temp indentation")

def creating_new_excel_from_filtered_column(filtered_column_array,VCentre_name):
	workbook2=xlwt.Workbook("encoding=utf-8")
	sheet1=workbook2.add_sheet('Sheet1',cell_overwrite_ok=True)
	#green_colour =xlwt.easyxf(strg_to_parse="pattern : pattern solid ,back_colour colour red")
	#green_colour =xlwt.easyxf('pattern : pattern_back_colour green')
	green_colour = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour green')
	red_colour = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour red')
	grey_colour = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour gray_ega')
	yellow_colour = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;pattern: pattern solid, fore_colour yellow')

	#border = xlwt.easyxf('font: bold off, color black; \ borders: left thin, right thin, top thin, bottom thin; pattern: pattern solid, fore_color black;')
	border_heading = xlwt.easyxf('font: bold off, color black;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thick, right thick, top thick, bottom thick;\
                     pattern: pattern solid, fore_color white;')
	border = xlwt.easyxf('borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin')
	#border = xlwt.easyxf('font: bold off, color black;\
    #                borders: top_color black, bottom_color black, right_color black, left_color black,\
    #                          left thick, right thick, top thick, bottom thick;\
    #                 pattern: pattern solid, fore_color aqua;')
	i=0 #row number
	#vmname#hwversion
	sheet1.write(0,0,filtered_column_array[0][0],border_heading)#title
	sheet1.write(0,1,filtered_column_array[2][0],border_heading)#title
	sheet1.write(0,2,filtered_column_array[4][0],border_heading)#title
	sheet1.write(0,3,filtered_column_array[1][0],border_heading)
	sheet1.write(0,4,filtered_column_array[3][0],border_heading)
	sheet1.write(0,5,"VCentre Name",border_heading)
	#sheet1.write(0,4,filtered_column_array[5][0])

	i=1  #row number
	for x in range(1, (len(filtered_column_array[2])-1) ):		
	
		if int(filtered_column_array[2][x])==13:
			sheet1.write(x,0,filtered_column_array[0][x],border)#vmname
			sheet1.write(x,1,filtered_column_array[2][x],green_colour)#hwversion
			sheet1.write(x,4,filtered_column_array[3][x],border)#os
			sheet1.write(x,5,VCentre_name,border)#vcentre name
		elif int(filtered_column_array[2][x]) < 13:
			sheet1.write(x,0,filtered_column_array[0][x],border)
			sheet1.write(x,1,filtered_column_array[2][x],red_colour)
			sheet1.write(x,4,filtered_column_array[3][x],border)#os
			sheet1.write(x,5,VCentre_name,border)#vcentre name
		elif int(filtered_column_array[2][x])>13:
			sheet1.write(x,0,filtered_column_array[0][x],border)
			sheet1.write(x,1,"Error : The version is greater than 13")
			sheet1.write(x,4,filtered_column_array[3][x],border)#os
			sheet1.write(x,5,VCentre_name,border)#vcentre name

		else:
			sheet1.write(x,0,filtered_column_array[0][x],border)
			sheet1.write(i,1,"Error : Not Defined")
			sheet1.write(0,4,filtered_column_array[3][x],border)
			sheet1.write(x,5,VCentre_name,border)#vcentre name
		#i=i+1

	#tools
		if filtered_column_array[4][x].lower() == "toolsok":
			sheet1.write(x,2,filtered_column_array[4][x],green_colour)
		elif filtered_column_array[4][x].lower() == "toolsnotrunning" :
			sheet1.write(x,2,filtered_column_array[4][x],yellow_colour)
		elif filtered_column_array[4][x].lower() == "toolsold":
			sheet1.write(x,2,filtered_column_array[4][x],grey_colour)
		elif filtered_column_array[4][x].lower() == "toolsnotinstalled":
			sheet1.write(x,2,filtered_column_array[4][x],red_colour)
		else:
			sheet1.write(x,2,filtered_column_array[4][x])


		if filtered_column_array[1][x].lower() == "green":
			sheet1.write(x,3,filtered_column_array[1][x],green_colour)
		elif filtered_column_array[1][x].lower() == "gray":
			sheet1.write(x,3,filtered_column_array[1][x],grey_colour)
		elif filtered_column_array[1][x].lower() == "yellow":
			sheet1.write(x,3,filtered_column_array[1][x],yellow_colour)
		elif filtered_column_array[1][x].lower() == "red": 
			sheet1.write(x,3,filtered_column_array[1][x],red_colour)
		else :
			sheet1.write(x,3,filtered_column_array[1][x])


	workbook2.save('temp.xls')

def vcentre_name_fetcher(path):
	path_part_array = path.split('\\')
	length = len(path_part_array)
	VCentre_name = path_part_array[length-1].replace(".xlsx","")

	return VCentre_name
	print(path_part_array[length-1].replace(".xlsx",""))
	print(path_part_array[length-1])
	print(type(path_part_array))
	print("\n\n",path_part_array)
	

def print_first_row_from__fetched_sheet(sheet):
	print("\n\n\n","Test : type of sheet:", type(sheet),"\n\n\n")
	print(sheet.row_values(0))

def tkinter_init():
	root = tkinter.Tk()
	root.geometry('800x800+250+200')
	return root
'''	
def tkinter_table_init(root):
	table = tktable.Table(root,state='disabled',width=50,titlerows=1,rows=5,cols=4,colwidth=20)
	#table = tktable.Table()
	#not working
	print("temp for indented block error")
'''
#def tkinter_create_table_manually(root):
	#print("temp for indented block error")

#working
def tkinter_ttk_Treeview_table_test(root):
	tree = tkinter.ttk.Treeview(master=root)
	tree.pack()
	tree.insert("", 0, iid=None,text="linux",tags="colorred")
	tree.tag_configure("colorred",background="red")
	tree.insert("", 0, iid=None,text="aws")
	#tree.insert("", index, iid=None,)

	
	#tree.grid(row=4,column=0,columnspan=2)
	#tree.heading('#0',text = 'Name')
	#tree.heading(2, text='Price')
	#tree.insert('',0, text = "123", values = "sdasdad")

#def tkinter_ttk_Treeview_table_with_excel(root,)

def pie_chart(size_array):
	print("temp indent")
	fig = pl.figure(figsize=(3, 3))
	labels = 'Ver 13','Other','Ver 10','Ver 8','Ver 7'
	sizes =size_array
	colours = ['yellowgreen','grey','gold','lightskyblue','red']
	explode = [0,0,0,0,0]
	pl.pie(sizes,explode = explode,labels = labels,colors = colours,autopct = '%1.1f%%',shadow =True,startangle = 140)
	pl.title('Page One')
	#patches, texts = pl.pie(sizes, colors=colours, shadow=True, startangle=90)
	#pl.legend(patches,labels,loc = "best")
	pl.axis('equal')
	pl.show()

	fig.savefig("graph.pdf",bbox_inches = 'tight')

def size_array_for_pie_chart(filtered_column_array):
	thirteen = 0
	ten =0
	eight =0
	seven = 0 
	other = 0
	
	for x in range(1,len(filtered_column_array[2])):
		if  int(filtered_column_array[2][x]) == 13:
			thirteen = thirteen + 1
		elif int(filtered_column_array[2][x]) ==10:
			ten =ten +1
		elif int(filtered_column_array[2][x]) ==8:
			eight =eight +1
		elif int(filtered_column_array[2][x]) ==7 :
			seven =seven +1
		else:
			other =1
	size_array = [thirteen,other,ten,eight,seven]
	return size_array





if __name__ == "__main__":
	path =path()
	if path !=1:	
		book=open_workbook(path)
		print("indentation temp")
	else:
		print("Wrong Path")


#from here we can put for loop
	#tempsheet =fetch_sheet_by_index(0,book)
	VCentre_name = vcentre_name_fetcher(path)
	
	sheet_array = []
	sheet_array=passing_sheet_name_to_fetch_sheet_by_name(book)
	print("\n\n","Test : func(main) : type of sheet_array",type(sheet_array),"\n\n",)
	print("\n\n\n",type(sheet_array),"\n\n\n")
	print_first_row_from__fetched_sheet(sheet_array[0])
	filtered_column_array = filtering_required_columns_from_fetched_sheets(sheet_array)
	creating_new_excel_from_filtered_column(filtered_column_array,VCentre_name)
	#root = tkinter_init()
	size_array = size_array_for_pie_chart(filtered_column_array)
	pie_chart(size_array)
	#table = tkinter_table_init(root) tktable issue
	#temp commenttkinter_ttk_Treeview_table_test(root)
	#root.mainloop()
	










