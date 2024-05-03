main_file = open('question.txt', encoding="utf-8")
all_lines = main_file.readlines()
converted_dict = {}
question = ''
correct_answer = ''
list_option = []
order = ''
counter = 1
for line in all_lines:
    
    if '#' in line:
        order = line[:-1]
        #print('order - ', line)
    elif line == '\n' or line == '':
        continue
    if '*!' in line:
        question = line[2:]
        list_option = []
        correct_answer = ''
        counter = 0
        #print(question)
    elif '*' in line:
        if '*+' in line:
            line.strip()
            if '++' in line:
                correct_answer = line[2:]
            elif '+' in line:
                correct_answer = line[1:]
            #print('option - ', correct_answer)
        else:
            list_option.append(line[1:])
            counter += 1
            #print('option - ', line[1:])
    
    if len(list_option) == 4:
        list_option.insert(0, correct_answer)
        converted_dict[order + ')' + question] = list_option
        question = ''
        correct_answer = ''
        list_option = []
        counter = 0
#print(converted_dict)
### end reading file
### start writin as template
'''
main_file.close()
template = open('template_2.txt', 'w+', encoding="utf-8")

for k,v in converted_dict.items():
    template.write(k)
    template.write('Бір ғана жауап :)\n')
    for elem in v:
        template.write(elem)
    template.write('\n')
template.close()

print(converted_dict)
'''


import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo_2.xlsx')
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

