import xlsxwriter
import sys
import re
#Checking file
if len(sys.argv) < 2:
        print ("Cannot find param file")        
        sys.exit(1)

sourceFile = sys.argv[1]
file = open(sourceFile, encoding="utf8")
targetFile = re.sub(".txt$", ".xlsx", sourceFile)
quesnumber = 1
row = 1
column = 1
index = 0
workbook = xlsxwriter.Workbook(targetFile)
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'STT')
worksheet.write(0, 1, 'Câu hỏi')
worksheet.write(0, 2, 'ĐA Đúng')
worksheet.write(0, 3, 'ĐA Sai 1')
worksheet.write(0, 4, 'ĐA Sai 2')
worksheet.write(0, 5, 'ĐA Sai 3')
worksheet.write(1, 0, quesnumber)
arr = ['','','','','']
def insert(arr, row):
    print("Waiting convert question: [" + str(quesnumber)+ "]")
    worksheet.write(row, 0, quesnumber)
    worksheet.write(row, 1, arr[0])
    worksheet.write(row, 2, arr[1])
    worksheet.write(row, 3, arr[2])
    worksheet.write(row, 4, arr[3])
    worksheet.write(row, 5, arr[4])

for line in file:
    line = line.replace("A) ", "")
    line = line.replace("B) ", "")
    line = line.replace("C) ", "")
    line = line.replace("D) ", "")
    line = line.lstrip(" ./,\"")
    line = line.rstrip(" /,\"")
    if "ANSWER:" in line:
        line = line[8:9]
        temp = arr[1]
        trueansw = ord(line) - 64 #ord('A') = 65 then subtract 64 get 1
        arr[1] = arr[trueansw] 
        arr[trueansw] = temp
        insert(arr,row)
        row+=1
        index=0
        quesnumber+=1
        arr = ['','','','','']
    else:
        if (index > 4):
            print("Question number " +  str(quesnumber) + " have too much answer!")
            print("Question content: "+arr[0])
            sys.exit(1)
        else:
            arr[index]=line
            index+=1

print("Convert sucess!!!")
workbook.close()
