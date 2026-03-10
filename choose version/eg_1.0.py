import os
import random
import subprocess
import sys
import time

import openpyxl as op


def build_sheet_map(workbook):
    return {str(i + 1): name for i, name in enumerate(workbook.sheetnames)}


def clear_screen():
    command = 'cls' if os.name == 'nt' else 'clear'
    subprocess.run(command, shell=True, check=False)


def intro():
    sys.stdout.write('loading')
    sys.stdout.flush()
    for _ in range(4):
        time.sleep(1)
        sys.stdout.write('.')
        sys.stdout.flush()
    time.sleep(1)
    clear_screen()


def load_sheet_data(workbook, sheet_name):
    sheet = workbook[sheet_name]

    voca = []
    chinese = []

    for cell in sheet['A'][1:]:
        if cell.value is not None:
            voca.append(str(cell.value).strip())

    for cell in sheet['B'][1:]:
        if cell.value is not None:
            chinese.append(str(cell.value).strip())

    size = min(len(voca), len(chinese))
    return voca[:size], chinese[:size]


def ask_sheet_key(sheet_map):
    while True:
        print('choose sheet')
        menu = ' '.join([f'{key} {name}' for key, name in sheet_map.items()])
        print(menu)
        key = input().strip()
        if key in sheet_map:
            return key
        print('wrong number')
        time.sleep(1)
        clear_screen()


def ask_test_type():
    while True:
        test_type = input('1 for voca trans chi, 2 for chi trans voca\n').strip()
        if test_type in ('1', '2'):
            return test_type
        print('wrong enter value, enter again')


def ask_test_number(max_count):
    while True:
        raw = input('input tests number\n').strip()
        try:
            count = int(raw)
        except ValueError:
            print('please input a number')
            continue

        if 1 <= count <= max_count:
            return count
        print('number out of range (1~{})'.format(max_count))


def build_options(correct_answer, answer_pool):
    wrong_pool = [item for item in answer_pool if item != correct_answer]
    wrong_count = min(3, len(wrong_pool))
    options = random.sample(wrong_pool, wrong_count)
    options.append(correct_answer)
    random.shuffle(options)
    return options


def ask_choice(question_str, options):
    while True:
        ans = input('ans:\n').strip()
        try:
            ans_num = int(ans)
        except ValueError:
            print('illegal input type')
            time.sleep(1)
            clear_screen()
            print(question_str)
            continue

        if 1 <= ans_num <= len(options):
            return ans_num - 1

        print('illegal input type')
        time.sleep(1)
        clear_screen()
        print(question_str)


def run_quiz(test_type, questions, voca_to_chi, chi_to_voca, voca, chinese):
    wrong_choices = []
    wrong_questions = []
    correct_answers = []

    order = list(range(len(questions)))
    random.shuffle(order)

    for index, shuffled_idx in enumerate(order, start=1):
        question = questions[shuffled_idx]

        if test_type == '1':
            correct = voca_to_chi[question]
            options = build_options(correct, chinese)
        else:
            correct = chi_to_voca[question]
            options = build_options(correct, voca)

        question_str = ""
        question_str += f"{index}: {question}\n"
        for i, option in enumerate(options, start=1):
            question_str += f"{i}. {option}\n"
        print(question_str)

        selected_idx = ask_choice(question_str, options)
        selected = options[selected_idx]

        if selected == correct:
            print('PASS')
        else:
            print('Wrong answer')
            wrong_choices.append(selected)
            wrong_questions.append(question)
            correct_answers.append(correct)

        time.sleep(1)
        clear_screen()

    return wrong_questions, wrong_choices, correct_answers

def statistics(wrong_questions, wrong_choices, correct_answers, total_questions):
    print('總共錯誤題數 :', len(wrong_questions))
    for i in range(len(wrong_questions)):
        print(
            '題目 : {0} ,你的選擇 : {1} , 正確答案 : {2}'.format(
                wrong_questions[i],
                wrong_choices[i],
                correct_answers[i],
            )
        )
    if total_questions == 0:
        print('正確率 : 0.00%')
        return

    correct_rate = ((total_questions - len(wrong_questions)) / total_questions) * 100
    print('正確率 : {:.2f}%'.format(correct_rate))


def main():
    clear_screen()
    intro()

    path = os.path.join(sys.path[0], 'en.xlsx')
    workbook = op.load_workbook(path)
    sheet_map = build_sheet_map(workbook)

    if not sheet_map:
        print('No worksheets found in Excel file.')
        return

    sheet_key = ask_sheet_key(sheet_map)
    voca, chinese = load_sheet_data(workbook, sheet_map[sheet_key])

    if not voca or not chinese:
        print('No available vocabulary data.')
        return

    voca_to_chi = {voca[i]: chinese[i] for i in range(len(voca))}
    chi_to_voca = {value: key for key, value in voca_to_chi.items()}

    clear_screen()
    print('avaliable problems :', len(voca))

    test_type = ask_test_type()
    test_number = ask_test_number(len(voca))

    if test_type == '1':
        questions = random.sample(voca, test_number)
    else:
        questions = random.sample(chinese, test_number)

    clear_screen()
    print('test start')
    time.sleep(2)
    clear_screen()

    wrong_questions, wrong_choices, correct_answers = run_quiz(
        test_type,
        questions,
        voca_to_chi,
        chi_to_voca,
        voca,
        chinese,
    )

    statistics(wrong_questions, wrong_choices, correct_answers, len(questions))

if __name__ == '__main__':
    main()