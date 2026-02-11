#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Обработчик HTML таблиц результатов соревнований.
Перетащите HTML файлы на exe (Windows) или на .app (macOS) для обработки.
"""

import sys
import os
import re
import platform
import subprocess
import traceback
import threading
import time
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from collections import defaultdict


# --- Определяем режим работы ---
IS_MACOS = platform.system() == 'Darwin'
IS_FROZEN = getattr(sys, 'frozen', False)
IS_WINDOWED = IS_FROZEN and IS_MACOS

# Файл лога для отладки (на рабочем столе)
if IS_MACOS:
    LOG_PATH = os.path.expanduser("~/Desktop/html_to_xlsx_log.txt")
else:
    LOG_PATH = None


def log(message):
    """Пишет в лог-файл и в stdout (если доступен)."""
    try:
        print(message)
    except Exception:
        pass
    if LOG_PATH:
        try:
            with open(LOG_PATH, 'a', encoding='utf-8') as f:
                f.write(f"{datetime.now().strftime('%H:%M:%S')} {message}\n")
        except Exception:
            pass


def notify(title, message):
    """Показать macOS уведомление или вывести в консоль."""
    if IS_MACOS:
        try:
            safe_msg = message.replace('"', '\\"').replace("'", "\\'")
            safe_title = title.replace('"', '\\"').replace("'", "\\'")
            subprocess.run([
                'osascript', '-e',
                f'display notification "{safe_msg}" with title "{safe_title}"'
            ], timeout=5)
        except Exception:
            pass
    log(f"[{title}] {message}")


def notify_error(message):
    """Показать уведомление об ошибке."""
    notify("html_to_xlsx — Ошибка", message)


def wait_before_exit():
    if IS_WINDOWED:
        return
    try:
        input("\nНажмите Enter для выхода...")
    except EOFError:
        pass


def get_output_folder(filepaths):
    """Определяет папку для сохранения результатов."""
    timestamp = datetime.now().strftime("%d.%m.%y %H-%M")
    folder_name = f"общая таблица {timestamp}"
    
    # Попытка 1: рядом с исходными файлами
    first_file_dir = os.path.dirname(os.path.abspath(filepaths[0]))
    output_folder = os.path.join(first_file_dir, folder_name)
    try:
        os.makedirs(output_folder, exist_ok=True)
        test_file = os.path.join(output_folder, ".test_write")
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        return output_folder, timestamp
    except (OSError, PermissionError):
        log(f"Нет прав записи в {first_file_dir}, пробуем Рабочий стол")
    
    # Попытка 2: Рабочий стол
    if IS_MACOS:
        desktop = os.path.expanduser("~/Desktop")
    else:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    
    output_folder = os.path.join(desktop, folder_name)
    try:
        os.makedirs(output_folder, exist_ok=True)
        return output_folder, timestamp
    except (OSError, PermissionError):
        log(f"Нет прав записи на Рабочий стол")
    
    # Попытка 3: домашняя папка
    output_folder = os.path.join(os.path.expanduser("~"), folder_name)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder, timestamp


def parse_html_file(filepath):
    """Парсит HTML файл и возвращает список (команда, место)."""
    results = []
    
    with open(filepath, 'rb') as f:
        raw_data = f.read()
    
    header = raw_data[:1000].decode('ascii', errors='ignore').lower()
    
    encoding = None
    if 'charset=windows-1251' in header or 'charset=cp1251' in header:
        encoding = 'cp1251'
    elif 'charset=utf-8' in header:
        encoding = 'utf-8'
    elif 'charset=koi8-r' in header:
        encoding = 'koi8-r'
    
    encodings_to_try = []
    if encoding:
        encodings_to_try.append(encoding)
    encodings_to_try.extend(['cp1251', 'utf-8', 'koi8-r', 'latin-1'])
    
    content = None
    for enc in encodings_to_try:
        try:
            content = raw_data.decode(enc)
            if content.count('\ufffd') < 10:
                break
        except (UnicodeDecodeError, LookupError):
            continue
    
    if content is None:
        log(f"Ошибка: не удалось прочитать файл {filepath}")
        return results
    
    soup = BeautifulSoup(content, 'html.parser')
    tables = soup.find_all('table')
    
    for table in tables:
        header_row = table.find('tr', bgcolor='silver')
        if not header_row:
            continue
        
        rows = table.find_all('tr')
        data_row = None
        for row in rows:
            cells = row.find_all('td')
            if cells and len(cells) >= 10:
                first_cell = cells[0].get_text(strip=True)
                if first_cell == '1':
                    data_row = row
                    break
        
        if not data_row:
            continue
        
        cells = data_row.find_all('td')
        
        step = 10
        for i in range(8, min(15, len(cells))):
            cell_text = cells[i].get_text(strip=True)
            if cell_text == '2':
                step = i
                break
        
        i = 0
        while i < len(cells):
            remaining_cells = len(cells) - i
            
            if remaining_cells < 4:
                break
            
            current_block_size = min(step, remaining_cells)
            participant_cells = cells[i:i+current_block_size]
            
            first = participant_cells[0].get_text(strip=True)
            if not first.isdigit():
                i += step
                continue
            
            team = participant_cells[3].get_text(strip=True)
            
            place = None
            for j in range(len(participant_cells) - 1, 3, -1):
                cell_text = participant_cells[j].get_text(strip=True)
                if cell_text and ':' not in cell_text:
                    if cell_text.isdigit() and len(cell_text) == 4:
                        year = int(cell_text)
                        if 1900 <= year <= 2100:
                            continue
                    
                    if cell_text.isdigit():
                        place = cell_text
                        break
                    cell_lower = cell_text.lower()
                    if any(x in cell_lower for x in ['н/ф', 'в/к', 'дск', 'снят', 'снт', 'дисквал']):
                        place = 'Сошел'
                        break
                    if 'пїЅ' in cell_text:
                        place = 'Сошел'
                        break
            
            if place is None and team:
                place = 'Сошел'
            
            if team and place:
                results.append((team, place))
            
            i += step
    
    return results


def extract_sort_key(team_name):
    match = re.match(r'^(\d+)\.\s*(.+)$', team_name)
    if match:
        return (int(match.group(1)), match.group(2).lower())
    return (float('inf'), team_name.lower())


def process_files(filepaths):
    """Обрабатывает все файлы и группирует результаты по командам."""
    teams_data = defaultdict(list)
    
    for filepath in filepaths:
        if not os.path.exists(filepath):
            log(f"Файл не найден: {filepath}")
            continue
        
        log(f"Обработка: {os.path.basename(filepath)}")
        results = parse_html_file(filepath)
        
        for team, place in results:
            teams_data[team].append(place)
    
    return teams_data


def create_xlsx(teams_data, output_path):
    """Создаёт xlsx файл с результатами."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты"
    
    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    headers = ["Команда", "Кол-во участников"] + [str(i) for i in range(1, 21)]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    sorted_teams = sorted(teams_data.keys(), key=extract_sort_key)
    
    for row_idx, team in enumerate(sorted_teams, 2):
        places = teams_data[team]
        
        cell = ws.cell(row=row_idx, column=1, value=team)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        cell = ws.cell(row=row_idx, column=2, value=len(places))
        cell.border = thin_border
        cell.alignment = cell_alignment
        
        for place_idx, place in enumerate(places[:20]):
            cell = ws.cell(row=row_idx, column=3 + place_idx, value=place)
            cell.border = thin_border
            cell.alignment = cell_alignment
        
        for empty_idx in range(len(places), 20):
            cell = ws.cell(row=row_idx, column=3 + empty_idx, value="")
            cell.border = thin_border
    
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 15
    for col_letter in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 
                       'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']:
        ws.column_dimensions[col_letter].width = 6
    
    wb.save(output_path)
    log(f"Файл сохранён: {output_path}")


def open_folder(path):
    """Открывает папку в Finder/Проводнике."""
    try:
        if IS_MACOS:
            subprocess.run(['open', path], timeout=5)
        elif platform.system() == 'Windows':
            os.startfile(path)
    except Exception:
        pass


def run_processing(filepaths):
    """Основная логика обработки файлов."""
    log(f"Получено файлов для обработки: {len(filepaths)}")
    for fp in filepaths:
        log(f"  -> {fp}")
    
    teams_data = process_files(filepaths)
    
    if not teams_data:
        notify_error("Не удалось извлечь данные из файлов!")
        return
    
    total_participants = sum(len(places) for places in teams_data.values())
    log(f"Найдено команд: {len(teams_data)}, участников: {total_participants}")
    
    output_folder, timestamp = get_output_folder(filepaths)
    output_filename = f"Результаты по командам {timestamp}.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    
    create_xlsx(teams_data, output_path)
    
    notify("html_to_xlsx — Готово!",
           f"Команд: {len(teams_data)}, участников: {total_participants}")
    
    open_folder(output_folder)
    log(f"Результаты сохранены в: {output_folder}")


def main_cli():
    """Запуск из командной строки (Windows или терминал macOS)."""
    if len(sys.argv) < 2:
        print("=" * 50)
        print("Обработчик HTML таблиц результатов соревнований")
        print("=" * 50)
        print("\nИспользование: перетащите HTML файлы на exe")
        print("или запустите: python html_to_xlsx_v2.py файл1.html файл2.html ...")
        wait_before_exit()
        return
    
    filepaths = sys.argv[1:]
    run_processing(filepaths)
    wait_before_exit()


def main_macos_app():
    """Запуск как macOS .app — получаем файлы через Apple Events."""
    try:
        from Foundation import NSObject
        from AppKit import NSApplication, NSApp
        
        class AppDelegate(NSObject):
            """Делегат приложения для обработки Apple Events (drag & drop)."""
            
            _files = []
            _processed = False
            
            def applicationWillFinishLaunching_(self, notification):
                log("AppDelegate: applicationWillFinishLaunching")
            
            def applicationDidFinishLaunching_(self, notification):
                log("AppDelegate: applicationDidFinishLaunching")
                # Даём время на получение файлов через Apple Events
                threading.Timer(1.0, self.checkAndProcess).start()
            
            def application_openFiles_(self, app, filenames):
                """Вызывается macOS при drag & drop файлов на .app."""
                log(f"AppDelegate: получены файлы через openFiles: {list(filenames)}")
                self._files.extend(filenames)
            
            def application_openFile_(self, app, filename):
                """Вызывается macOS при открытии одного файла."""
                log(f"AppDelegate: получен файл через openFile: {filename}")
                self._files.append(filename)
                return True
            
            def checkAndProcess(self):
                """Проверяет наличие файлов и запускает обработку."""
                if self._processed:
                    return
                self._processed = True
                
                filepaths = list(self._files)
                
                # Если файлы не пришли через Apple Events, проверяем sys.argv
                if not filepaths:
                    argv_files = [
                        arg for arg in sys.argv[1:]
                        if os.path.isfile(arg)
                    ]
                    if argv_files:
                        filepaths = argv_files
                        log(f"Файлы из sys.argv: {filepaths}")
                
                if filepaths:
                    try:
                        run_processing(filepaths)
                    except Exception as e:
                        log(f"Ошибка обработки: {e}")
                        log(traceback.format_exc())
                        notify_error(str(e))
                else:
                    notify("html_to_xlsx", "Перетащите HTML файлы на иконку приложения")
                    log("Нет входных файлов")
                
                # Завершаем приложение
                NSApp.terminate_(None)
        
        app = NSApplication.sharedApplication()
        delegate = AppDelegate.alloc().init()
        app.setDelegate_(delegate)
        
        log("Запуск NSApplication run loop...")
        app.run()
        
    except ImportError as e:
        log(f"PyObjC не доступен: {e}")
        log("Fallback на CLI режим")
        main_cli()


def main():
    """Точка входа."""
    try:
        # Очищаем лог при каждом запуске
        if LOG_PATH:
            try:
                with open(LOG_PATH, 'w', encoding='utf-8') as f:
                    f.write(f"=== html_to_xlsx запуск {datetime.now()} ===\n")
                    f.write(f"sys.argv: {sys.argv}\n")
                    f.write(f"platform: {platform.system()} {platform.machine()}\n")
                    f.write(f"IS_WINDOWED: {IS_WINDOWED}\n")
                    f.write(f"IS_FROZEN: {IS_FROZEN}\n\n")
            except Exception:
                pass
        
        if IS_WINDOWED:
            # macOS .app — нужен NSApplication для Apple Events
            main_macos_app()
        else:
            # Windows / терминал — файлы через sys.argv
            main_cli()
        
    except Exception as e:
        error_msg = f"Критическая ошибка: {str(e)}"
        log(error_msg)
        log(traceback.format_exc())
        notify_error(str(e))
        wait_before_exit()


if __name__ == "__main__":
    main()
