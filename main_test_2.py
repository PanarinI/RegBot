import os
import pandas as pd
from datetime import datetime
import re


def check_file_exists(file_path):
    if os.path.exists(file_path):
        print(f"Файл '{file_path}' найден.")
        return True
    else:
        print(f"Файл '{file_path}' не найден. Завершение программы.")
        return False


def load_file(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Ошибка при загрузке файла: {e}")
        return None


def save_file(df, file_path):
    try:
        df.to_excel(file_path, index=False)
        print("Изменения сохранены.")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")


def add_new_entry(df):
    new_entry = {}
    columns = df.columns.tolist()

    for column in columns:
        if column == "Файл размещен в папке?":
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
        elif column == "Ссылка на OneDrive (input)":
            while True:
                value = input(f"Введите значение для '{column}' (ссылка должна содержать 'iframe src='): ").strip()
                if "iframe src=" in value:
                    new_entry[column] = value
                    break
                else:
                    print("Ошибка: ссылка должна содержать 'iframe src='. Проверьте и введите снова.")
        elif column == "Ссылка на OneDrive (обработанная)":
            input_link = new_entry.get("Ссылка на OneDrive (input)", "")
            if input_link:
                processed_link = re.sub(r'width="[0-9]+"', 'width="90%"', input_link)
                processed_link = re.sub(r'height="[^"]+"', 'height="1800"', processed_link)
                processed_link = f"<p align=\"center\">{processed_link}</p>"
                new_entry[column] = processed_link
            else:
                new_entry[column] = ""
        elif column == "Дата публикации":
            while True:
                value = input(f"Введите значение для '{column}' (формат YYYY-MM-DD): ").strip()
                try:
                    new_entry[column] = datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
                    break
                except ValueError:
                    print("Некорректный формат даты. Используйте формат 'YYYY-MM-DD'.")
        else:
            value = input(f"Введите значение для '{column}': ").strip()
            new_entry[column] = value

    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    print("Запись добавлена.")
    return df


def modify_table_structure(df):
    print("\nТекущие столбцы таблицы:")
    print(", ".join(df.columns.tolist()))

    print("\nВыберите действие:")
    print("1. Добавить новый столбец")
    print("2. Удалить существующий столбец")
    print("3. Переименовать столбец")

    action = input("Введите номер действия (1/2/3): ").strip()

    if action == "1":
        new_column = input("Введите название нового столбца: ").strip()
        if new_column in df.columns:
            print("Столбец уже существует.")
        else:
            df[new_column] = ""
            print(f"Столбец '{new_column}' добавлен.")
    elif action == "2":
        del_column = input("Введите название столбца для удаления: ").strip()
        if del_column in df.columns:
            df.drop(columns=[del_column], inplace=True)
            print(f"Столбец '{del_column}' удалён.")
        else:
            print("Столбец не найден.")
    elif action == "3":
        old_column = input("Введите название столбца для переименования: ").strip()
        if old_column in df.columns:
            new_column = input("Введите новое название столбца: ").strip()
            if new_column in df.columns:
                print("Столбец с таким именем уже существует.")
            else:
                df.rename(columns={old_column: new_column}, inplace=True)
                print(f"Столбец '{old_column}' переименован в '{new_column}'.")
        else:
            print("Столбец не найден.")
    else:
        print("Некорректный выбор.")

    return df


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
            if column == "Ссылка на OneDrive (обработанная)":
                continue  # Пропускаем, т.к. она рассчитывается автоматически
            current_value = df.at[index, column]
            new_value = input(f"Введите новое значение для '{column}' (оставьте пустым для сохранения '{current_value}'): ").strip()
            if new_value:
                if column == "Ссылка на OneDrive (input)" and "iframe src=" not in new_value:
                    print("Ошибка: ссылка должна содержать 'iframe src='. Проверьте и введите снова.")
                    continue
                if column == "Дата публикации":
                    try:
                        new_value = datetime.strptime(new_value, "%Y-%m-%d").strftime("%Y-%m-%d")
                    except ValueError:
                        print("Некорректный формат даты. Используйте формат 'YYYY-MM-DD'.")
                        continue
                df.at[index, column] = new_value

        # Автоматически обновляем обработанную ссылку
        input_link = df.at[index, "Ссылка на OneDrive (input)"]
        if input_link:
            processed_link = re.sub(r'width="[0-9]+"', 'width="90%"', input_link)
            processed_link = re.sub(r'height="[^"]+"', 'height="1800"', processed_link)
            processed_link = f"<p align=\"center\">{processed_link}</p>"
            df.at[index, "Ссылка на OneDrive (обработанная)"] = processed_link

        print("Запись обновлена.")
    except ValueError:
        print("Некорректный ввод.")
    return df


def main():
    file_path = "database_test2.xlsx"

    if not check_file_exists(file_path):
        return

    df = load_file(file_path)
    if df is None:
        return

    while True:
        print("\nМеню:")
        print("1. Внести новую запись")
        print("2. Отредактировать имеющуюся запись")
        print("3. Добавить/изменить структуру таблицы")
        print("4. Выйти")

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
            print("Выход из программы.")
            break
        else:
            print("Некорректный выбор. Попробуйте снова.")


if __name__ == "__main__":
    main()
