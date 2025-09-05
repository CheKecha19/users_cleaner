# processors/kontur_processor.py
import pandas as pd
from utils import load_kontur_data, find_duplicates, find_internal_duplicates, find_users_to_remove

def process_kontur_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из Контура"""
    if 3 not in selected_options and 0 not in selected_options:
        return df, {}
    
    print("Обработка данных Контура...")
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        print("Предупреждение: AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_kontur': 0,
            'internal_duplicates_kontur': 0,
            'users_to_remove_kontur': pd.DataFrame()
        }
    
    # Загружаем данные из Контура
    kontur_data = load_kontur_data()
    
    if not kontur_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        kontur_fio = kontur_data['Контур_ФИО'][:len(df)]
        kontur_admin = kontur_data['Контур_Администратор'][:len(df)]
        kontur_status = kontur_data['Контур_статус'][:len(df)]  # Используем новое название
        
        df['Контур_ФИО'] = pd.Series(kontur_fio)
        df['Контур_Администратор'] = pd.Series(kontur_admin)
        df['Контур_статус'] = pd.Series(kontur_status)  # Используем новое название
    
    # Разделение на отдельные DataFrame (используем новое название столбца)
    kontur_df = df[['Контур_ФИО', 'Контур_статус']].dropna(subset=['Контур_ФИО'])
    
    # Инициализация результатов
    results = {
        'duplicates_ad_kontur': 0,
        'internal_duplicates_kontur': 0,
        'users_to_remove_kontur': pd.DataFrame()
    }
    
    # Поиск дубликатов для Контура
    results['duplicates_ad_kontur'] = len(find_duplicates(ad_employees_df, kontur_df, 'AD_ФИО', 'Контур_ФИО'))
    results['internal_duplicates_kontur'] = len(find_internal_duplicates(kontur_df, 'Контур_ФИО'))
    results['users_to_remove_kontur'] = find_users_to_remove(kontur_df, ad_employees_df, ad_employees_df)
    
    return df, results