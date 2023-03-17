'''
some function names are deprecated -> new function names
documentInfo -> metadata
getNumPgaes() -> len(a.pages)
getPage() -> a.pages[]
extractText() -> extract_text()

'''
def conversion(s):
    new = ""                         # function to convert array of character into string
    for x in s:
        new += x
    return new

import PyPDF2
import xlsxwriter
import hashlib                       # library for hash_value to remove duplicacy
import os

print('NOTE: Make sure Pdf is located in this directory')
name = input('enter excel sheet name you want to create : ')     # taking excel file name form user
workbook = xlsxwriter.Workbook(name+'.xlsx')       # made a workbook instance with name of xls sheet
worksheet = workbook.add_worksheet('1st')                     # made [worksheet] as object and accessed add_worksheet method
worksheet.write(0,0,'Order No.')                         # to write in 1st row & 1st column
pdf = input('enter Pdf name from where to extract order no. ')
a = PyPDF2.PdfReader(f'{pdf}.pdf')                          # opening pdf to read
page = len(a.pages)     # a is always used as we have made it as object
print(f'no of pages -> {page}')             # printing of number of pages
str = ""                # empty string
arr = []                # for storing all words
with open('ExtractedOrderno.txt','w') as f:           # to reduce redundancy of order no's
    f.truncate(0)
    f.close()
for i in range(0,page):     #index will start from 0 that is equals to first page of pdf
    str += a.pages[i].extract_text()    # to iterate to page and extract text at same time

    with open("ExtractedText.txt",'w') as f:    # opening of txt file to write all the text of PDF
        f.write(str)
        f.close()
    f = open("ExtractedText.txt","r")
    content = f.read()
    for word in content.split():       # file is read word by word   # understand the split function
        arr.append(word)             # all words are appended in list   # array is made to store strings
    for j in range(len(arr)):            # visisting array of string element by element
        arr_final = [char for char in arr[j]]         # for converting string to char-array  and storing if arr_final
        if arr_final[len(arr_final)-1]=='1' and arr_final[len(arr_final)-2]=='_':    # to check for the pattern of order no.
            resultant_data = conversion(arr_final)                # converting that finded array of character to string
            with open('ExtractedOrderno.txt','a+') as f:
                f.write(resultant_data)                        # writing order no to text file # but here comes the redundant values
                f.write('\n')
        else:
            pass
    del arr[:]              # array of all words is deleted because if not done than arr of string keep on increasing

lines_present = set()
output_file = open('DataToInsert.txt','w')
for l in open('ExtractedOrderno.txt','r'):
    hash_value = hashlib.md5(l.rstrip().encode('utf-8')).hexdigest()
    if hash_value not in lines_present:                                 # code to remove duplicacy in order no.
        output_file.write(l)
        lines_present.add(hash_value)
output_file.close()

with open('DataToInsert.txt','r') as f:
    anss = f.readlines()                                    # file from which data is to be entered
    f.close()
for i in range(len(anss)):
    print(f'Order No.-> {anss[i]}',end='')

worksheet.write_column('A2',anss)                                # writing data in column
workbook.close()
print(f"Order no's are succesfuly entred in Excel file - {name}")
os.system(f"{name}.xlsx")
