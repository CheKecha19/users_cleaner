# processors/onec_processor.py
import pandas as pd
from utils import load_onec_data, find_duplicates, find_internal_duplicates, find_users_to_remove

def process_onec_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из 1С"""
    if 1 not in selected_options and 0 not in selected_options:
        return df, {}
    
    print("Обработка данных 1С...")
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        print("Предупреждение: AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_1c': 0,
            'internal_duplicates_1c': 0,
            'users_to_remove_1c': pd.DataFrame()
        }
    
    # Загружаем данные из 1С
    onec_data = load_onec_data()
    
    if not onec_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        onec_fio = onec_data['1C_ФИО'][:len(df)]
        onec_active = onec_data['1C_Активен'][:len(df)]
        
        df['1C_ФИО'] = pd.Series(onec_fio)
        df['1C_Активен'] = pd.Series(onec_active)
    
    # Разделение на отдельные DataFrame
    onec_df = df[['1C_ФИО', '1C_Активен']].dropna(subset=['1C_ФИО'])
    
    # Инициализация результатов
    results = {
        'duplicates_ad_1c': 0,
        'internal_duplicates_1c': 0,
        'users_to_remove_1c': pd.DataFrame()
    }
    
    # Поиск дубликатов для 1С
    results['duplicates_ad_1c'] = len(find_duplicates(ad_employees_df, onec_df, 'AD_ФИО', '1C_ФИО'))
    results['internal_duplicates_1c'] = len(find_internal_duplicates(onec_df, '1C_ФИО'))
    results['users_to_remove_1c'] = find_users_to_remove(onec_df, ad_employees_df, ad_employees_df)
    
    return df, results