import pandas as pd
from openpyxl import Workbook

# Функция для загрузки данных из файла Excel (CSV)
def load_data(file_path):
    """
    Загружает данные расписания из CSV-файла.
    """
    return pd.read_csv(file_path, delimiter=';', encoding='utf-8')

# Функция для фильтрации данных
def filter_schedule(data, filter_by=None, filter_value=None):
    """
    Фильтрует данные расписания по заданному критерию.
    """
    if filter_by and filter_value:
        # Проверяем, есть ли такой фильтр в данных
        if filter_by in data.columns:
            return data[data[filter_by].str.strip() == filter_value.strip()]
        else:
            print(f"Ошибка: Колонка '{filter_by}' не найдена.")
    return data

# Функция для экспорта данных в новый Excel-файл
def export_to_excel(data, output_file='schedule.xlsx'):
    """
    Экспортирует данные расписания в Excel-файл.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Расписание"

    # Записываем заголовки
    ws.append(list(data.columns))

    # Записываем строки данных
    for _, row in data.iterrows():
        ws.append(list(row))

    # Сохраняем файл
    wb.save(output_file)
    print(f"Расписание сохранено в файл: {output_file}")

# Главная функция программы
def main():
    # Укажите путь к вашему файлу CSV
    file_path = "Книга1_extended.csv"  # Замените на путь к вашему файлу
    output_file = "filtered_schedule.xlsx"

    # Загружаем данные
    data = load_data(file_path)
    print("Данные загружены успешно.")

    # Выбор фильтра
    print("\nВыберите критерий фильтрации расписания:")
    print("1. Для группы")
    print("2. Для факультета")
    print("3. Для всего вуза")
    choice = input("Введите номер варианта (1/2/3): ").strip()

    filter_by = None
    filter_value = None

    if choice == '1':
        filter_by = 'Группа'
        filter_value = input("Введите название группы (например, 'Группа А'): ").strip()
    elif choice == '2':
        filter_by = 'Факультет'
        filter_value = input("Введите название факультета (например, 'Факультет 1'): ").strip()

    # Фильтруем данные
    filtered_data = filter_schedule(data, filter_by, filter_value)

    # Проверка на пустой результат
    if filtered_data.empty:
        print("Нет данных, соответствующих выбранным критериям.")
    else:
        # Экспортируем данные в Excel
        export_to_excel(filtered_data, output_file=output_file)

if __name__ == "__main__":
    main()
