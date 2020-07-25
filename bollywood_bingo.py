from datetime import date
from googletrans import Translator
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from tabulate import tabulate
import random as r
import time

def board_generation(name, date, seed_number):
    r.seed(seed_number)

    actor_array = [
        'Abhishek Bachchan', 'Sonam Kapoor', 'Deepika Padukone', 'Ranveer Singh', 'Imran Khan',
        'Tiger Shroff', 'Shah Rukh Khan', 'Emraan Hashmi', 'Irrfan Khan', 'Boman Irani',
        'Aamir Khan', 'Hrithik Roshan', 'Saif Ali Khan', 'Alia Bhatt', 'Sara Ali Khan',
        'Amitabh Bachchan', 'Anupam Kher', 'Raj Kapoor', 'Kader Khan', 'Johnny Lever',
        'Geeta Bali', 'Shyama', 'Asha Parekh', 'Rajendra Kumar', 'Dev Anand'
    ]
    zero_arr = [0 for x in range(len(actor_array))]

    header = [name, '', '', '', date]
    bingo = 'BINGO'
    bingo_board = [header, list(bingo)]

    col_length = 5
    col_arr = []
    while actor_array != zero_arr:
        if col_length == 0:
            bingo_board.append(col_arr)
            col_length = 5
            col_arr = []

        loc = r.randint(0, len(actor_array) - 1)
        if actor_array[loc] != 0:
            col_arr.append(actor_array[loc])
            actor_array[loc] = 0
            col_length -= 1

        if actor_array == zero_arr:
            bingo_board.append(col_arr)

    return bingo_board


def random_pick(all_entries, seed_number):
    r.seed(seed_number)
    zero_arr = [0 for x in range(len(all_entries))]
    while all_entries != zero_arr:
        loc = r.randint(0, len(all_entries) - 1)
        if all_entries[loc] != 0:
            print(all_entries[loc])
            input("Press Enter to continue...\n\n")
            all_entries[loc] = 0
        else:
            continue


def write_to_excel(workbook, bingo_board, sheet_number, font_size=18):
    # Set styles
    side = Side(border_style='thin')
    font = Font(name='Arial', size=font_size)
    border = Border(
        left=side,
        right=side,
        top=side,
        bottom=side)
    alignment = Alignment(horizontal="center", vertical="center")
    worksheet = workbook.create_sheet('0{}'.format(sheet_number))
    for row_index, row in enumerate(bingo_board):
        # Append row to worksheet
        worksheet.append(row)
        # Format row height
        corrected_index = row_index + 1
        if corrected_index == 1:
            worksheet.row_dimensions[corrected_index].height = 40
        else:
            worksheet.row_dimensions[corrected_index].height = 60

    for col in 'ABCDE':
        worksheet.column_dimensions[col].width = 35
    for row in worksheet["1:7"]:
        for cell in row:
            cell.font = font
            cell.border = border
            cell.alignment = alignment

    return workbook


def translate_board(bingo_board, source, dest):
    translator = Translator()
    new_board = bingo_board[0:2]

    for row in bingo_board[2:]:
        translations = translator.translate(row, src=source, dest=dest)
        new_board.append(x.text for x in translations)

    return new_board


def main():
    seed = 1
    sheet_number = 1
    names = ['P1', 'P2', 'P3', 'P4', 'P5', 'P6']
    today = date.today()
    all_entries = []
    workbook = Workbook()

    for name in names:
        # Generate bingo board for each person
        bingo_board = board_generation(name, today.strftime("%m/%d/%Y"), seed)
        # Provide translation of board in different language
        if name == 'P5' or name == 'P6':
            translated_board = translate_board(bingo_board, 'en', 'gu')
            workbook = write_to_excel(workbook, translated_board, sheet_number, font_size=32)
        else:
            workbook = write_to_excel(workbook, bingo_board, sheet_number)
        # Print fancy grid for board
        print(tabulate(bingo_board, tablefmt="fancy_grid"))

        # Increase seed to generate new random board
        seed += 1
        # Increase worksheet number
        sheet_number += 1
        # Remove first 2 rows w/ extra information
        split_bingo_board = bingo_board[2:]
        # Generate list of all possible values & add to single array
        for arr in split_bingo_board:
            B = 'B - {}'.format(arr[0])
            I = 'I - {}'.format(arr[1])
            N = 'N - {}'.format(arr[2])
            G = 'G - {}'.format(arr[3])
            O = 'O - {}'.format(arr[4])
            all_entries.extend([B, I, N, G, O])
    print('\n--------------------------')
    print('Starting the BINGO Round!!')
    print('--------------------------\n')
    # Save the workbook
    workbook.remove(workbook['Sheet'])
    workbook.save("bollywood_bingo.xlsx")
    # Use a different seed for picking locations
    seed = 10
    # random_pick(all_entries, seed)

if __name__ == '__main__':
    main()
