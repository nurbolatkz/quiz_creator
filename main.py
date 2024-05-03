import openpyxl

class Quiz:
    def __init__(self):
        self.questions = []

    def read_questions_from_file(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            question = None
            for line in lines:
                line = line.strip()
                if line.startswith('!'):
                    if question:
                        self.questions.append(question)
                    question = {'Question Text': line[2:].strip(), 'Question Type': 'Multiple Choice', 'Options': [], 'Correct Answer': None}
                
                elif line.startswith('*') or line.startswith('•') :
                    option = line[1:].strip()
                    if option.startswith('+') :
                        question['Correct Answer'] = option[1:]
                        #print(line)
                        question['Options'].append(option[1:])
                    else:
                        #print(line)
                        question['Options'].append(option)
            if question:
                self.questions.append(question)

    def write_to_excel(self, excel_file):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['Question Text', 'Question Type', 'Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Correct Answer', 'Time in seconds', 'Image Link'])
        counter = 0;
        for question in self.questions:
            counter += 1
            print(counter)
            options = question['Options'] + [''] * (5 - len(question['Options']))  # Pad options if less than 5
            #print(question['Options'])
            correct_answer_index = question['Options'].index(question['Correct Answer'])  # Get index of correct answer
            sheet.append([question['Question Text'], question['Question Type']] + options + [correct_answer_index+1, '', ''])
        
        workbook.save(excel_file)
        print("Finished writing to file")

# Example usage
quiz = Quiz()
quiz.read_questions_from_file('quiz_questions.txt')
quiz.write_to_excel('quiz_questions.xlsx')
