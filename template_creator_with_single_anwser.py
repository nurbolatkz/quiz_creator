temp = open('formula.txt', encoding="utf-8")
            
question = []
data = {}

lines = temp.readlines()

correct_option = ''
q = ''


for line in lines:
    if len(line)< 1:
        print(line)
        continue
    elif '*+' in line:
        correct_option = line[:-2]
        data[q] = correct_option
        q = ''
        correct_option = ''
    else:
        q = line[:-2]

import random  
    
values = list(data.values())

for el in data.keys():
    data[el] = [data[el]]
    for i in range(4):
        while(True):
            a = random.choice(values)
            if a in data[el]:
                continue
            else:
                data[el].append(a)
                break
  
converted_dict = data
import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo_formula.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Question Text')
worksheet.write('B1', 'Question Type')
worksheet.write('C1', 'Option 1')
worksheet.write('D1', 'Option 2')
worksheet.write('E1', 'Option 3')
worksheet.write('F1', 'Option 4')
worksheet.write('G1', 'Option 5')
worksheet.write('H1', 'Correct Answer')
worksheet.write('I1', 'Time in seconds')
worksheet.write('J1', 'Image Link')
#Multiple Choice

# Text with formatting.
#worksheet.write('A2', 'World', bold)
rowIndex = 2
for k,v in converted_dict.items():
    worksheet.write('A' + str(rowIndex), k)
    worksheet.write('B' + str(rowIndex), 'Multiple Choice')
    counter = 2
    #print('write ', rowIndex)
    for elem in v:
        #print(len(v))
        if elem == '' or elem == '\n':
            continue
        worksheet.write(rowIndex-1, counter, elem)
        counter += 1
    
    worksheet.write(rowIndex-1, 7, 1)
    worksheet.write(rowIndex-1, 8, 35)
    counter = 0
    rowIndex += 1
    

workbook.close()
print('Finished')
