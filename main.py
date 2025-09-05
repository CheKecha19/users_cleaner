# main.py
import logging
import pandas as pd
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILE
from excel_processor import process_excel_data
from ad_export import export_ad_users

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(OUTPUT_DIR / "processing.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def get_user_choice():
    """Получение выбора пользователя"""
    print("\n" + "="*50)
    print("Выберите опции для проверки (через пробел):")
    print("0 - Всё")
    print("1 - 1С")
    print("2 - Сфера (Диадок)")
    print("3 - Контур")
    print("="*50)
    
    while True:
        choice = input("Ваш выбор: ").strip()
        
        if not choice:
            print("Пожалуйста, введите хотя бы одну цифру")
            continue
            
        choices = choice.split()
        
        # Проверка на валидность ввода
        valid_choices = {'0', '1', '2', '3'}
        if all(c in valid_choices for c in choices):
            # Если выбран 0, добавляем все остальные опции
            if '0' in choices:
                return {0, 1, 2, 3}
            return set(int(c) for c in choices)
        else:
            print("Некорректный ввод. Пожалуйста, используйте цифры 0, 1, 2, 3 через пробел")

def get_employee_type_choice():
    """Получение выбора типа сотрудников"""
    print("\n" + "="*50)
    print("Выберите тип сотрудников для проверки (через пробел):")
    print("0 - Все")
    print("1 - Сотрудники")
    print("2 - ГПХ")
    print("="*50)
    
    while True:
        choice = input("Ваш выбор: ").strip()
        
        if not choice:
            print("Пожалуйста, введите хотя бы одну цифру")
            continue
            
        choices = choice.split()
        
        # Проверка на валидность ввода
        valid_choices = {'0', '1', '2'}
        if all(c in valid_choices for c in choices):
            # Если выбран 0, добавляем все остальные опции
            if '0' in choices:
                return {0, 1, 2}
            return set(int(c) for c in choices)
        else:
            print("Некорректный ввод. Пожалуйста, используйте цифры 0, 1, 2 через пробел")

def main():
    logging.info("Запуск обработки данных")
    
    # Получаем выбор пользователя
    selected_options = get_user_choice()
    selected_employee_types = get_employee_type_choice()
    logging.info(f"Выбранные опции: {selected_options}")
    logging.info(f"Выбранные типы сотрудников: {selected_employee_types}")
    
    # Экспорт данных из AD (всегда выполняется)
    try:
        logging.info("Экспорт пользователей из Active Directory")
        total_users, employees_count, gph_count = export_ad_users()
        logging.info(f"Экспорт AD завершен: {total_users} пользователей, {employees_count} сотрудников, {gph_count} ГПХ")
    except Exception as e:
        logging.error(f"Ошибка при экспорте из AD: {e}")
        logging.info("Продолжение обработки с пустыми данными AD")
        total_users, employees_count, gph_count = 0, 0, 0
    
    # Обработка Excel данных
    try:
        logging.info("Обработка Excel данных")
        results = process_excel_data(selected_options, selected_employee_types)
        
        logging.info("Обработка завершена. Результаты:")
        if 1 in selected_options or 0 in selected_options:
            logging.info(f"- Дубликаты между AD и 1С: {results.get('duplicates_ad_1c', 0)}")
            logging.info(f"- Внутренние дубликаты в 1С: {results.get('internal_duplicates_1c', 0)}")
            logging.info(f"- Пользователей для удаления из 1С: {len(results.get('users_to_remove_1c', pd.DataFrame()))}")
        if 2 in selected_options or 0 in selected_options:
            logging.info(f"- Дубликаты между AD и Диадок: {results.get('duplicates_ad_diadoc', 0)}")
            logging.info(f"- Внутренние дубликаты в Диадоке: {results.get('internal_duplicates_diadoc', 0)}")
            logging.info(f"- Пользователей для удаления из Диадока: {len(results.get('users_to_remove_diadoc', pd.DataFrame()))}")
        if 3 in selected_options or 0 in selected_options:
            logging.info(f"- Дубликаты между AD и Контур: {results.get('duplicates_ad_kontur', 0)}")
            logging.info(f"- Внутренние дубликаты в Контуре: {results.get('internal_duplicates_kontur', 0)}")
            logging.info(f"- Пользователей для удаления из Контура: {len(results.get('users_to_remove_kontur', pd.DataFrame()))}")
        logging.info(f"- Несоответствий между AD и Штатным расписанием: {results.get('comparison_count', 0)}")
        
    except Exception as e:
        logging.error(f"Ошибка при обработке Excel: {str(e)}")
    
    logging.info(f"Результаты сохранены в файл: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()