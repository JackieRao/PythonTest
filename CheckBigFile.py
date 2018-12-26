import os, openpyxl

#The function check the size
def checkbigfile(folder,size):
	folder = os.path.abspath(folder)
	largesize = size
	for foldername, subfoldernames, filenames in os.walk(folder):
		for filename in filenames:
			filepath = os.path.join(foldername,filename)
			if os.path.getsize(filepath) >= largesize:
				yield filename, filepath, os.path.getsize(filepath)
				
#The function save the big file list to excel file			
def save_excel(files):
	wb = openpyxl.Workbook()
	sheet = wb.active
	sheet.title = 'Big_File_List'
	rows = 1
	sheet['A'+str(rows)].value = "FileName"
	sheet['B'+str(rows)].value = "FilePath"
	sheet['C'+str(rows)].value = "FileSize(MB)"
	for filename, filepath, filesize in files:
		rows += 1
		sheet['A'+str(rows)].value = filename
		sheet['B'+str(rows)].value = filepath
		sheet['C'+str(rows)].value = round(filesize/1024/1024,2)
	wb.save('Big_File_List.xlsx')

if __name__=="__main__":
    #Validate the input value is number
	while True:				
		checkfolder = input("Please input a folder path: ")
		if os.path.exists(checkfolder):
			break
		else:
			print ("The folder doesn't exist, please input the correct folder path.")
	#Check if the Unit is correct
	while True:				
		size = input("Please input threshold size: ")
		size_num = 0
	
		if size.upper().endswith('GB'):
			if size[:-2].isdigit():
				size_num = int(size[:-2])*1024**3
				break
			else:
				print ("Please input number!")
		elif size.upper().endswith('MB'):
			if size[:-2].isdigit():
				size_num = int(size[:-2])*1024**2
				break
			else:
				print ("Please input number!")
		elif size.upper().endswith('KB'):
			if size[:-2].isdigit():
				size_num = int(size[:-2])*1024
				break
			else:
				print ("Please input number!")
		else:
			print ("Please input a correct Unit like GB, MB, KB")
	files = checkbigfile(checkfolder, size_num)
	#Delete the excel file when it exist
	if os.path.exists('Big_File_List.xlsx'):
		os.remove('Big_File_List.xlsx')
	save_excel(files)
	