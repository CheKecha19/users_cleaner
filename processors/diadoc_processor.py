# processors/diadoc_processor.py
import pandas as pd
from utils import load_diadoc_data, find_duplicates, find_internal_duplicates, find_users_to_remove

def process_diadoc_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из Диадока"""
    if 2 not in selected_options and 0 not in selected_options:
        return df, {}
    
    print("Обработка данных Диадока...")
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        print("Предупреждение: AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_diadoc': 0,
            'internal_duplicates_diadoc': 0,
            'users_to_remove_diadoc': pd.DataFrame()
        }
    
    # Загружаем данные из Диадока
    diadoc_data = load_diadoc_data()
    
    if not diadoc_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        diadoc_fio = diadoc_data['Диадок_ФИО'][:len(df)]
        diadoc_active = diadoc_data['Диадок_Активен'][:len(df)]
        diadoc_admin = diadoc_data['Диадок_Администратор'][:len(df)]
        
        df['Диадок_ФИО'] = pd.Series(diadoc_fio)
        df['Диадок_Активен'] = pd.Series(diadoc_active)
        df['Диадок_Администратор'] = pd.Series(diadoc_admin)
    
    # Разделение на отдельные DataFrame
    diadoc_df = df[['Диадок_ФИО', 'Диадок_Активен']].dropna(subset=['Диадок_ФИО'])
    
    # Инициализация результатов
    results = {
        'duplicates_ad_diadoc': 0,
        'internal_duplicates_diadoc': 0,
        'users_to_remove_diadoc': pd.DataFrame()
    }
    
    # Поиск дубликатов для Диадока
    results['duplicates_ad_diadoc'] = len(find_duplicates(ad_employees_df, diadoc_df, 'AD_ФИО', 'Диадок_ФИО'))
    results['internal_duplicates_diadoc'] = len(find_internal_duplicates(diadoc_df, 'Диадок_ФИО'))
    results['users_to_remove_diadoc'] = find_users_to_remove(diadoc_df, ad_employees_df, ad_employees_df)
    
    return df, results