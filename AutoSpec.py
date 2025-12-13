import os
import re
import json
import shutil

from openpyxl import load_workbook
from colorama import Fore, Style, init
import win32com.client as win32

init(autoreset=True)

SPEC_SHEET_NAME = "Specification"
CONFIG_FILE = "config.json"


# ===================== ЦВЕТА =====================

def yellow(text): return Fore.YELLOW + text + Style.RESET_ALL
def green(text): return Fore.GREEN + text + Style.RESET_ALL
def red(text): return Fore.LIGHTRED_EX + text + Style.RESET_ALL


# ===================== КОНФИГ =====================

def script_dir():
    return os.path.dirname(os.path.abspath(__file__))


def load_config():
    path = os.path.join(script_dir(), CONFIG_FILE)
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(data):
    path = os.path.join(script_dir(), CONFIG_FILE)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ===================== ВСПОМОГАТЕЛЬНЫЕ =====================

def normalize_path(path):
    return path.strip().strip('"').strip("'")


def parse_ranges(user_input):
    result = set()
    cleaned = user_input.replace(" ", "").replace(";", ",")
    for part in cleaned.split(","):
        if "-" in part:
            a, b = part.split("-")
            result.update(range(int(a), int(b) + 1))
        elif part:
            result.add(int(part))
    return result


def extract_invoice_number(folder):
    m = re.match(r"(\d+)", folder)
    return int(m.group(1)) if m else None


def hide_all_except_spec_xlsx(path):
    wb = load_workbook(path)
    if SPEC_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f'Лист "{SPEC_SHEET_NAME}" не найден')

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if sheet_name == SPEC_SHEET_NAME:
            ws.sheet_state = "visible"
        else:
            ws.sheet_state = "hidden"

    wb.save(path)


# ===================== СЦЕНАРИИ =====================

def scenario_xlsx(folder, number):
    src = os.path.join(folder, f"Invoice {number}.xlsx")
    dst = os.path.join(folder, f"Invoice {number} fcs.xlsx")
    shutil.copy2(src, dst)
    hide_all_except_spec_xlsx(dst)


def scenario_xls(folder, number):
    src = os.path.join(folder, f"Invoice {number}.xlsx")
    dst = os.path.join(folder, f"Invoice {number} fcs.xls")

    excel = win32.Dispatch("Excel.Application")
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(src)
        wb.SaveAs(dst, FileFormat=56)
        wb.Close(SaveChanges=False)

        wb = excel.Workbooks.Open(dst)
        for sheet in wb.Worksheets:
            sheet.Visible = sheet.Name == SPEC_SHEET_NAME
        wb.Save()
        wb.Close()
    finally:
        excel.Quit()


# ===================== MAIN =====================

def main():
    config = load_config()
    base_dir = config.get("base_dir")

    while True:
        if not base_dir:
            print(yellow("Укажите путь к директории с инвойсами:"))
            base_dir = normalize_path(input("> "))
            if not os.path.isdir(base_dir):
                print(red("Некорректный путь"))
                base_dir = None
                continue
            save_config({"base_dir": base_dir})

        print(yellow("\nВыберите сценарий:"))
        print(yellow("1) Создание файла спецификации XLS") + green("X"))
        print(yellow("2) Создание файла спецификации XLS"))
        current_path = base_dir if base_dir else "не указан"
        print(yellow(f'3) Изменить путь к директории с инвойсами (Сейчас: {current_path})'))
        print(yellow("0) Выход"))

        choice = input("> ").strip()

        if choice == "0":
            break

        if choice == "3":
            base_dir = None
            continue

        if choice not in {"1", "2"}:
            print(red("Неверный выбор"))
            continue

        while True:
            print(yellow("\nУкажите диапазон номеров инвойсов:"))
            invoice_numbers = parse_ranges(input("> "))

            matched = []
            for folder in os.listdir(base_dir):
                num = extract_invoice_number(folder)
                if num in invoice_numbers:
                    matched.append((num, os.path.join(base_dir, folder)))

            if not matched:
                print(red("Инвойсы указанных номеров не были найдены"))
                continue

            processed = 0
            for number, folder in sorted(matched):
                print(yellow(f"[В РАБОТЕ] Invoice {number}"))
                try:
                    if choice == "1":
                        scenario_xlsx(folder, number)
                    else:
                        scenario_xls(folder, number)

                    processed += 1
                    print(green(f"[ГОТОВО] Invoice {number}"))
                except Exception as e:
                    print(red(f"[ОШИБКА] Invoice {number}: {e}"))

            print(green(f"\nГотово! Обработано файлов: {processed}"))
            break  # назад в меню сценариев

    print(green("\nВыход из утилиты"))
    input(yellow("Нажмите Enter для закрытия окна"))


if __name__ == "__main__":
    main()
