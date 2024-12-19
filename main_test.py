import os
import pandas as pd
from datetime import datetime

def check_file_exists(file_path):
    if os.path.exists(file_path):
        print(f"Файл '{file_path}' найден.")
        return True
    else:
        print(f"Файл '{file_path}' не найден.")
        return False


def load_or_create_file(file_path):
    if not os.path.exists(file_path):
        print("Создание нового файла...")
        columns = [
            "Имя материала", "Полная аннотация", "Краткая аннотация", "Дата публикации",
            "Раздел", "Файл размещен в папке?", "Ссылка на ЯДиск", "Ссылка на OneDrive",
            "Ссылка на OneDrive (обработанная)", "Название страницы для Tilda"
        ]
        df = pd.DataFrame(columns=columns)
        df.to_excel(file_path, index=False)
    else:
        df = pd.read_excel(file_path)
    return df

def save_file(df, file_path):
    df.to_excel(file_path, index=False)
    print("Изменения сохранены.")


import re

def add_new_entry(df):
    new_entry = {}
    columns = df.columns.tolist()

    for column in columns:
        if column == "In_folder?":
            while True:
                value = input(f"Введите значение для '{column}' (да/нет): ").strip().lower()
                if value in ["да", "нет"]:
                    if value == "нет":
                        print("Поместите файл в нужную папку.")
                    else:
                        new_entry[column] = value
                        break
                else:
                    print("Некорректный ввод. Введите 'да' или 'нет'.")
        elif column == "text?":
            while True:
                value = input(f"Введите значение для '{column}' (да/нет): ").strip().lower()
                if value in ["да", "нет"]:
                    new_entry[column] = value
                    if value == "да":
                        new_entry["YaDisk_link/Rutube_link"] = input("Укажите ссылку на Яндекс.Диске: ").strip()
                    elif value == "нет":
                        new_entry["YaDisk_link/Rutube_link"] = input("Укажите ссылку на Rutube: ").strip()
                    break
                else:
                    print("Некорректный ввод. Введите 'да' или 'нет'.")
        elif column == "OneDrive_link (input)":
            while True:
                value = input(f"Введите значение для '{column}' (ссылка должна содержать 'iframe src='): ").strip()
                if "iframe src=" in value:
                    new_entry[column] = value
                    break
                else:
                    print("Ошибка: ссылка должна содержать 'iframe src='. Проверьте и введите снова.")

        elif column == "OneDrive_link (processed)":
            input_link = new_entry.get("OneDrive_link (input)", "")
            if input_link:
                # Отладочный вывод
                print(f"Original input_link: {input_link}")

                # Проверяем, что input_link — строка
                if not isinstance(input_link, str):
                    input_link = str(input_link)

                # Убираем старые значения width и height и заменяем их на новые
                processed_link = re.sub(r'width="[0-9]+"', 'width="90%"', input_link)
                print(f"After replacing width: {processed_link}")

                processed_link = re.sub(r'height="[^"]+"', 'height="1800"', processed_link)
                print(f"After replacing height: {processed_link}")

                # Добавляем обрамление
                processed_link = f"<p align=\"center\">{processed_link}</p>"
                print(f"Final processed link: {processed_link}")
                new_entry[column] = processed_link
            else:
                print("Warning: input_link is empty.")
                new_entry[column] = ""

        elif column == "Page_name_Tilda":
            value = input(f"Введите значение для '{column}': ").strip()
            new_entry[column] = value
        elif column not in ["YaDisk_link/Rutube_link"]:  # Пропускаем автоматизированные столбцы
            value = input(f"Введите значение для '{column}': ").strip()
            if column == "Дата публикации":
                try:
                    value = datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
                except ValueError:
                    print("Некорректный формат даты. Используйте формат 'YYYY-MM-DD'.")
                    return df
            new_entry[column] = value

    while True:
        # Вывод предварительных данных для подтверждения
        print("\nПредварительная запись:")
        for key, value in new_entry.items():
            print(f"{key}: {value}")

        # Запрос подтверждения
        confirm = input("Подтвердить добавление записи? (да/нет/отредактировать): ").strip().lower()
        if confirm == "да":
            # Добавление записи в таблицу
            new_row = pd.DataFrame([new_entry])  # Создаем DataFrame из одной записи
            df = pd.concat([df, new_row], ignore_index=True)  # Используем pd.concat
            print("Запись добавлена.")
            return df
        elif confirm == "нет":
            print("Добавление записи отменено.")
            return df
        elif confirm == "отредактировать":
            print("Редактирование записи:")
            for column in columns:
                if column in new_entry:
                    current_value = new_entry[column]
                    new_value = input(f"Введите новое значение для '{column}' (оставьте пустым для сохранения '{current_value}'): ").strip()
                    if new_value:
                        if column == "Дата публикации":
                            try:
                                new_value = datetime.strptime(new_value, "%Y-%m-%d").strftime("%Y-%m-%d")
                            except ValueError:
                                print("Некорректный формат даты. Используйте формат 'YYYY-MM-DD'.")
                                continue
                        new_entry[column] = new_value
            print("Изменения сохранены.")
        else:
            print("Некорректный ввод. Введите 'да', 'нет' или 'отредактировать'.")

def edit_entry(df):
    try:
        index = int(input("Введите номер записи для редактирования (0 для первой записи): "))
        if index < 0 or index >= len(df):
            print("Некорректный номер записи.")
            return df

        print("Текущая запись:")
        print(df.iloc[index])

        columns = df.columns.tolist()
        for column in columns:
            if column == "OneDrive_link (processed)":
                # Пропускаем обработку processed ссылки, программа сама пересчитает ее
                continue

            current_value = df.at[index, column]
            if column == "OneDrive_link (input)":
                # Обрабатываем `OneDrive_link (input)` с проверкой
                while True:
                    new_value = input(f"Введите новое значение для '{column}' (оставьте пустым для сохранения '{current_value}'): ").strip()
                    if not new_value:
                        # Если пользователь ничего не ввел, сохраняем текущее значение
                        break
                    if "iframe src=" in new_value:
                        df.at[index, column] = new_value
                        # Автоматически обновляем processed ссылку
                        processed_link = re.sub(r'width="[0-9]+"', 'width="90%"', new_value)
                        processed_link = re.sub(r'height="[^"]+"', 'height="1800"', processed_link)
                        processed_link = f"<p align=\"center\">{processed_link}</p>"
                        df.at[index, "OneDrive_link (processed)"] = processed_link
                        print("OneDrive_link (processed) автоматически обновлен.")
                        break
                    else:
                        print("Ошибка: ссылка должна содержать 'iframe src='. Проверьте и введите снова.")
            else:
                # Для остальных столбцов
                new_value = input(f"Введите новое значение для '{column}' (оставьте пустым для сохранения '{current_value}'): ").strip()
                if new_value:
                    df.at[index, column] = new_value

        print("Запись обновлена.")
    except ValueError:
        print("Некорректный ввод.")
    return df


def modify_table_structure(df):
    # Показываем текущие столбцы таблицы
    print("\nТекущие столбцы таблицы:")
    print(", ".join(df.columns.tolist()))  # Показываем только названия столбцов

    print("\nВыберите действие:")
    print("1. Добавить новый столбец")
    print("2. Удалить существующий столбец")
    print("3. Переименовать столбец")

    action = input("Введите номер действия (1/2/3): ").strip()

    if action == "1":
        # Добавить новый столбец
        print("\nТекущие столбцы таблицы:")
        print(", ".join(df.columns.tolist()))  # Список названий столбцов
        new_column = input("Введите название нового столбца: ").strip()
        if new_column in df.columns:
            print("Столбец уже существует.")
        else:
            df[new_column] = ""
            print(f"Столбец '{new_column}' добавлен.")

    elif action == "2":
        # Удалить существующий столбец
        print("\nТекущие столбцы таблицы:")
        print(", ".join(df.columns.tolist()))  # Список названий столбцов
        del_column = input("Введите название столбца для удаления: ").strip()
        if del_column in df.columns:
            df.drop(columns=[del_column], inplace=True)
            print(f"Столбец '{del_column}' удалён.")
        else:
            print("Ошибка: столбец не найден. Убедитесь в правильности ввода.")

    elif action == "3":
        # Переименовать столбец
        while True:
            print("\nТекущие столбцы таблицы:")
            print(", ".join(df.columns.tolist()))  # Список названий столбцов
            old_column = input(
                "Введите название столбца, который нужно переименовать (или введите 'отмена' для выхода): ").strip()
            if old_column.lower() == "отмена":
                print("Переименование столбца отменено.")
                break
            if old_column in df.columns:
                new_column = input(f"Введите новое название для столбца '{old_column}': ").strip()
                if new_column in df.columns:
                    print("Ошибка: столбец с таким именем уже существует.")
                else:
                    df.rename(columns={old_column: new_column}, inplace=True)
                    print(f"Столбец '{old_column}' переименован в '{new_column}'.")
                    break
            else:
                print("Ошибка: такого столбца не существует. Убедитесь в правильности ввода.")

    else:
        print("Некорректный выбор. Попробуйте снова.")

    return df


def main():
    file_path = "database_test2.xlsx"

    if not check_file_exists(file_path):
        create_new = input("Создать новый файл? (да/нет): ").lower()
        if create_new != "да":
            return

    df = load_or_create_file(file_path)

    while True:
        print("\nМеню:")
        print("1. Внести новую запись")
        print("2. Отредактировать имеющуюся запись")
        print("3. Добавить/изменить структуру таблицы")
        print("4. Открыть файл")
        print("5. Выйти")

        choice = input("Введите номер действия: ")

        if choice == "1":
            df = add_new_entry(df)
            save_file(df, file_path)
        elif choice == "2":
            df = edit_entry(df)
            save_file(df, file_path)
        elif choice == "3":
            df = modify_table_structure(df)
            save_file(df, file_path)
        elif choice == "4":
            open_file(file_path)
        elif choice == "5":
            print("Выход из программы.")
            break
        else:
            print("Некорректный выбор. Попробуйте снова.")


def open_file(file_path):
    os.system(f"start {file_path}")
    print("Файл открыт в стандартной программе.")

def main():
    file_path = "database_test2.xlsx"

    if not check_file_exists(file_path):
        create_new = input("Создать новый файл? (да/нет): ").lower()
        if create_new != "да":
            return

    df = load_or_create_file(file_path)

    while True:
        print("\nМеню:")
        print("1. Внести новую запись")
        print("2. Отредактировать имеющуюся запись")
        print("3. Добавить/изменить структуру таблицы")
        print("4. Открыть файл")
        print("5. Выйти")

        choice = input("Введите номер действия: ")

        if choice == "1":
            df = add_new_entry(df)
            save_file(df, file_path)
        elif choice == "2":
            df = edit_entry(df)
            save_file(df, file_path)
        elif choice == "3":
            df = modify_table_structure(df)
            save_file(df, file_path)
        elif choice == "4":
            open_file(file_path)
        elif choice == "5":
            print("Выход из программы.")
            break
        else:
            print("Некорректный выбор. Попробуйте снова.")

if __name__ == "__main__":
    main()
