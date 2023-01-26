import openpyxl

# Чтение файла ecxel и запрет на его редактирование
error_list_xlsx = openpyxl.open("stanok_error_list.xlsx", read_only=True)
cycle = True


# Функция выбора станка, которая ссылается на индекс страницы ecxel
def machine_selection():
    while True:
        ktf = ["ktf", "KTF", "КТФ", "ктф", "rna", "RNF", "ЛЕА", "леа"]
        ml = ["ML", "ml", "мл", "МЛ", "ьд", "ЬД", "vk", "VK", "TBT", "tbt", "тбт", "ТБТ", "n,n", "N,N", "ЕИЕ", "еие"]
        kss_800 = ["kss 800", "KSS 800", "ксс 800", "КСС 800", "800", "ЛЫЫ 800", "лыы 800", "rcc 800", "RCC 800"]
        kss_1250 = ["kss 1250", "KSS 1250", "ксс 1250", "КСС 1250", "1250", "ЛЫЫ 1250", "лыы 1250", "rcc 1250", "RCC 1250"]
        machine = input("Выберите ваш станок: ")  # Выбор станка по его названию
        if machine.lower() in ml:
            return error_list_xlsx.worksheets[0]
        elif machine.lower() in ktf:
            return error_list_xlsx.worksheets[1]
        elif machine.lower() in kss_800:
            return error_list_xlsx.worksheets[2]
        elif machine.lower() in kss_1250:
            return error_list_xlsx.worksheets[3]
        else:
            print("Станок не найден")


error_list = machine_selection()  # Эта переменная получает индекс страницы из функции machine_selection

# Если станок не найден, то цикл выполнения поиска ошибки не запускается

if error_list != error_list_xlsx.worksheets[0] and error_list != error_list_xlsx.worksheets[1] and error_list != error_list_xlsx.worksheets[2] and error_list != error_list_xlsx.worksheets[3]:
    cycle = False
else:
    cycle = True

while cycle:

    number = input("Введите номер вашей ошибки: ")


    # Функция нахождения ошибки по ее номеру
    def error_search(error_number):
        for i in range(1, error_list.max_row + 1):
            error_line = error_list[i][0].value
            if not str(error_number).isdigit():
                return "Это не номер ошибки"
            elif str(error_number) in error_line[:6] and str(error_number) == error_line[:6] and str(
                    error_number).isdigit() is True:
                error_text_0 = error_list[i][0].value  # Столбец с номером ошибки и ее содержанием
                error_text_1 = error_list[i][1].value  # Столбец с решением ошибки

                return f"Ваша ошибка:{error_text_0[6:]}\nРешение ✅: {error_text_1}"

        else:
            return "Такой ошибки не существует!"


    print()
    print(error_search(number))
    print()
