import itertools
from string import digits, punctuation, ascii_letters
import win32com.client as client
from datetime import datetime
import time


def excel_doc():
    print('==== Привет юный хакер ====')

    try:
        password_len = input("Напиши какой длины был пароль, от скольки - до скольки символов: ")
        password_len = [int(item) for item in password_len.split("-")]
    except:
        print('Проверьте введеные данные')

    print('Если пароль состоит только из цифр, введи: 1\nЕсли только из букв, введи: 2'
          'Если цифры и буквы, введи: 3\nЕсли цифры, буквы и символы, введи4: 4')

    try:
        choice = int(input('Твой выбор?:'))
        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            possible_symbols = 'Ничего не понял'
        print(possible_symbols)
    except:
        print('Ничего не понял')

    start_time = time.time()
    print(f"Мы начали в - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    count = 0

    for pass_len in range(password_len[0], password_len[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_len):
            password = ''.join(password)

            open_doc = client.Dispatch('Excel.Application')
            count += 1

            try:
                open_doc.Workbooks.Open(
                    r'C:\Users\Dell\PycharmProjects\pythonProject5\Probe.xlsx',
                    False,
                    True,
                    None,
                    password
                )
                time.sleep(0.1)
                print(f"Закончили в - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
                print(f"Вресмя на взлом - {time.time() - start_time}")

                return f"Количество попыток #{count} Ваш пароль: {password}"
            except:
                print(f"Количество попыток #{count} Неверные пароли {password}")

def main():
    excel_doc()


if __name__ == '__main__':
    main()