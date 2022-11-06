from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import random

def execute() -> None:
    NUMBER_OF_CATEGORIES = 62

    wb = load_workbook('aeae_categories.xlsx')
    ws = wb.worksheets[0]

    categories = []
    database_count = []
    questions_count = []

    for i in range(2, NUMBER_OF_CATEGORIES + 2):
        categories.append(ws[f'A{i}'].value)
        database_count.append(ws[f'B{i}'].value)
        questions_count.append(ws[f'C{i}'].value)

    random_questions_indices = []
    current_index = 1

    for required, total in zip(questions_count, database_count):
        indices = sorted(random.sample(
            range(current_index, current_index + total), required
        ))

        random_questions_indices += indices
        current_index += total

    print(random_questions_indices)
    random.shuffle(random_questions_indices)
    print(random_questions_indices)

    
if __name__ == '__main__':
    execute()
