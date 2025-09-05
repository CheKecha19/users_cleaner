# ad_export.py
import subprocess
import pandas as pd
import os
import logging
from tqdm import tqdm
import sys
import json
import unicodedata
from config import AD_EXPORT_DIR, OUTPUT_DIR

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(OUTPUT_DIR / "ad_export.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

def clean_value(value):
    """Очистка и преобразование значений"""
    if value is None:
        return ""
    
    # Преобразуем в строку
    cleaned = str(value)
    
    # Удаляем управляющие символы (0x00-0x1F) и спецсимволы Excel
    cleaned = ''.join(ch for ch in cleaned if unicodedata.category(ch)[0] != "C")
    cleaned = cleaned.replace('\x00', '').replace('\x01', '').replace('\x02', '')
    
    return cleaned.strip()

def export_ad_users():
    # Определяем путь для сохранения файлов
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Создаем директорию, если она не существует
    AD_EXPORT_DIR.mkdir(exist_ok=True)
    
    txt_filename = OUTPUT_DIR / 'ad_users_export.txt'
    xlsx_filename = OUTPUT_DIR / 'ad_users_export.xlsx'
    employees_filename = AD_EXPORT_DIR / 'сотрудники.txt'
    gph_filename = AD_EXPORT_DIR / 'ГПХ.txt'
    
    logging.info("="*60)
    logging.info("Начало экспорта пользователей Active Directory")
    logging.info(f"Файлы будут сохранены в: {script_dir}")
    logging.info(f"Разделенные файлы будут сохранены в: {AD_EXPORT_DIR}")
    
    # Исправленная PowerShell команда с добавлением DistinguishedName
    ps_command = """
    $OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $ErrorActionPreference = 'Stop'
    try {
        $users = Get-ADUser -Filter * -Properties Name, SamAccountName, Enabled, EmailAddress, Company, DistinguishedName
        $count = $users.Count
        Write-Host "Найдено пользователей: $count"
        
        # Обрабатываем каждого пользователя отдельно
        foreach ($user in $users) {
            $user | Select-Object Name, SamAccountName, Enabled, EmailAddress, Company, DistinguishedName |
            ConvertTo-Json -Depth 2 -Compress
            Write-Host ""  # Разделитель между записями
        }
    }
    catch {
        Write-Error $_
        exit 1
    }
    """
    
    logging.info("Запуск PowerShell команды...")
    try:
        # Запускаем PowerShell процесс
        process = subprocess.Popen(
            ["powershell", "-Command", ps_command],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1
        )
        
        # Переменные для сбора данных
        users = []
        user_count = 0
        current_json = ""
        in_json = False
        
        # Читаем вывод построчно
        logging.info("Обработка вывода PowerShell...")
        with tqdm(desc="Получение данных", unit="польз.") as pbar:
            while True:
                line = process.stdout.readline()
                if not line:  # Конец вывода
                    break
                    
                # Ищем количество пользователей
                if "Найдено пользователей:" in line:
                    try:
                        user_count = int(line.split(":")[1].strip())
                        logging.info(f"Найдено пользователей: {user_count}")
                        pbar.total = user_count
                    except:
                        pass
                    continue
                    
                # Пустые строки - разделители между JSON
                if line.strip() == "":
                    if current_json:
                        try:
                            user_data = json.loads(current_json)
                            users.append(user_data)
                            pbar.update(1)
                            current_json = ""
                            in_json = False
                        except json.JSONDecodeError:
                            logging.warning(f"Ошибка декодирования JSON: {current_json}")
                            current_json = ""
                    continue
                    
                # Собираем JSON строки
                current_json += line
                in_json = True
        
        # Проверяем завершающий JSON
        if current_json and in_json:
            try:
                user_data = json.loads(current_json)
                users.append(user_data)
                pbar.update(1)
            except json.JSONDecodeError:
                logging.warning(f"Ошибка декодирования последнего JSON: {current_json}")
        
        # Проверяем ошибки
        stderr = process.stderr.read()
        if stderr:
            logging.error(f"Ошибка PowerShell: {stderr}")
            if not users:
                # Создаем пустые файлы, если не удалось получить данные
                open(employees_filename, 'w', encoding='utf-8').close()
                open(gph_filename, 'w', encoding='utf-8').close()
                return 0, 0, 0
        
        if not users:
            logging.warning("Не найдено пользователей в Active Directory")
            # Создаем пустые файлы
            open(employees_filename, 'w', encoding='utf-8').close()
            open(gph_filename, 'w', encoding='utf-8').close()
            return 0, 0, 0
        
        # Обработка данных пользователей
        logging.info("Обработка данных...")
        processed_users = []
        employees = []  # Сотрудники кампуса
        gph_users = []  # Сотрудники ГПХ
        
        required_fields = ['Name', 'SamAccountName', 'Enabled', 'EmailAddress', 'Company', 'DistinguishedName']
        
        with tqdm(total=len(users), desc="Обработка данных", unit="польз.") as pbar:
            for user in users:
                processed_user = {}
                for field in required_fields:
                    value = user.get(field, "")
                    # Для поля Enabled сохраняем статус активности
                    if field == 'Enabled':
                        processed_user[field] = "Активна" if value else "Заблокирована"
                    else:
                        processed_user[field] = clean_value(value)
                
                processed_users.append(processed_user)
                
                # Разделение пользователей по критериям (только активные)
                dn = processed_user.get('DistinguishedName', '').lower()
                is_active = user.get('Enabled', False)
                
                if is_active:
                    # Сотрудники кампуса: DN содержит "cu_users" и не содержит "гпх"
                    if 'cu_users' in dn and 'гпх' not in dn:
                        employees.append(processed_user)
                    # Сотрудники ГПХ: DN содержит "external_organizations" или "гпх"
                    elif 'external_organizations' in dn or 'гпх' in dn:
                        gph_users.append(processed_user)
                
                pbar.update(1)
        
        # Экспорт в TXT (общий файл)
        logging.info(f"Экспорт в TXT файл: {txt_filename}")
        with open(txt_filename, 'w', encoding='utf-8') as txt_file, \
             tqdm(total=len(processed_users), desc="Запись в TXT", unit="польз.") as pbar:
            
            for user in processed_users:
                txt_file.write("=" * 80 + "\n")
                for key, value in user.items():
                    txt_file.write(f"{key}: {value}\n")
                txt_file.write("\n")
                pbar.update(1)
        
        # Экспорт сотрудников кампуса
        logging.info(f"Экспорт сотрудников кампуса: {employees_filename}")
        with open(employees_filename, 'w', encoding='utf-8') as emp_file, \
             tqdm(total=len(employees), desc="Запись сотрудников", unit="польз.") as pbar:
            
            for user in employees:
                emp_file.write(f"Name: {user['Name']}\n")
                emp_file.write(f"Status: {user['Enabled']}\n\n")
                pbar.update(1)
        
        # Экспорт сотрудников ГПХ
        logging.info(f"Экспорт сотрудников ГПХ: {gph_filename}")
        with open(gph_filename, 'w', encoding='utf-8') as gph_file, \
             tqdm(total=len(gph_users), desc="Запись ГПХ", unit="польз.") as pbar:
            
            for user in gph_users:
                gph_file.write(f"Name: {user['Name']}\n")
                gph_file.write(f"Status: {user['Enabled']}\n\n")
                pbar.update(1)
        
        # Экспорт в XLSX (общий файл)
        logging.info(f"Экспорт в XLSX файл: {xlsx_filename}")
        with tqdm(total=1, desc="Создание Excel", leave=False) as pbar:
            df = pd.DataFrame(processed_users)
            
            # Сохраняем в Excel
            df.to_excel(xlsx_filename, index=False, engine='openpyxl')
            pbar.update(1)
        
        logging.info("Экспорт завершен успешно!")
        logging.info(f"- TXT файл: {txt_filename}")
        logging.info(f"- Excel файл: {xlsx_filename}")
        logging.info(f"- Сотрудники кампуса: {employees_filename}")
        logging.info(f"- Сотрудники ГПХ: {gph_filename}")
        logging.info(f"- Всего экспортировано пользователей: {len(processed_users)}")
        logging.info(f"- Сотрудников кампуса: {len(employees)}")
        logging.info(f"- Сотрудников ГПХ: {len(gph_users)}")
        
        return len(processed_users), len(employees), len(gph_users)
    
    except Exception as e:
        logging.exception("Произошла критическая ошибка:")
        # Создаем пустые файлы при ошибке
        try:
            open(employees_filename, 'w', encoding='utf-8').close()
            open(gph_filename, 'w', encoding='utf-8').close()
        except:
            pass
        return 0, 0, 0

if __name__ == "__main__":
    export_ad_users()