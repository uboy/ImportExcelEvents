import pandas as pd
import re
import logging
import argparse
import os
import sys
import time
from datetime import datetime, time as dtime, timedelta
import win32com.client as win32
#from win32com.client import constants
import pytz

# === Константы ===
CATEGORY_NAME = "AutoImportSchedule"
DEFAULT_EXCEL_NAME = "Расписание для студентов.xlsx"

# === Настройка логирования ===
def setup_logger(log_file):
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Очистка обработчиков если перезапускаем main()
    if logger.hasHandlers():
        logger.handlers.clear()

    # Лог в файл
    fh = logging.FileHandler(log_file, encoding='utf-8')
    fh.setLevel(logging.INFO)
    fh.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    # Лог в консоль
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# === Парсинг времени из текста ===
def extract_times(text):
    if pd.isna(text):
        return None, None
    text = str(text).strip()

    # Формат "с 19.00 до 21.00"
    match = re.search(r"с\s*(\d{1,2})[.:](\d{2})?\s*до\s*(\d{1,2})[.:](\d{2})?", text)
    if match:
        return dtime(int(match.group(1)), int(match.group(2) or 0), 0), dtime(int(match.group(3)), int(match.group(4) or 0), 0)

    # Формат "18-21"
    match = re.search(r"(\d{1,2})\s*-\s*(\d{1,2})", text)
    if match:
        return dtime(int(match.group(1)), 0, 0), dtime(int(match.group(2)), 0, 0)

    return None, None


# === Загрузка участников ===
def load_invitees(invitees_file=None, invitees_arg=None):
    emails = set()

    if invitees_file and os.path.exists(invitees_file):
        with open(invitees_file, "r", encoding="utf-8") as f:
            for line in f:
                email = line.strip()
                if email:
                    emails.add(email)

    if invitees_arg:
        for email in invitees_arg.split(","):
            email = email.strip()
            if email:
                emails.add(email)

    return list(emails)


# === Загрузка Excel ===
def load_excel(excel_path):
    df = pd.read_excel(excel_path, sheet_name="Лист1")
    df = df.iloc[2:].reset_index(drop=True)
    return df


# === Преобразование в список событий ===
def parse_schedule(df):
    records = []
    for i in range(0, len(df), 2):
        if i + 1 >= len(df):
            break
        dates_row = df.iloc[i]
        events_row = df.iloc[i + 1]

        for col in df.columns:
            date_val = dates_row[col]
            event_val = events_row[col]

            if pd.isna(date_val) or pd.isna(event_val):
                continue

            # Парсим дату
            try:
                date = pd.to_datetime(date_val).date()
            except:
                continue

            # Парсим время
            start_time, end_time = extract_times(event_val)
            if not start_time or not end_time:
                continue

            # Убираем время из текста
            title = re.sub(r"(с\s*\d{1,2}[.:]\d{2}\s*до\s*\d{1,2}[.:]\d{2})|(\d{1,2}\s*-\s*\d{1,2})", "", str(event_val)).strip()

            records.append({
                "Date": date,
                "StartTime": start_time,
                "EndTime": end_time,
                "Title": title
            })
    return records


# === Подключение к Outlook ===
def connect_to_outlook(max_retries=3, retry_delay=5):
    """Подключение к Outlook с выбором аккаунта и папки календаря"""
    for attempt in range(1, max_retries + 1):
        try:
            logging.info(f"Попытка {attempt} подключения к Outlook...")
            outlook = win32.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace('MAPI')

            # Выбор аккаунта
            account = select_account(namespace)
            if not account:
                raise RuntimeError("Не удалось выбрать учетную запись Outlook.")

            # Выбор папки календаря
            calendar = select_calendar_folder(account)
            if not calendar:
                raise RuntimeError("Не удалось выбрать папку календаря.")

            logging.info("Успешное подключение к Outlook.")
            return outlook, namespace, calendar

        except Exception as e:
            logging.warning(f"Попытка {attempt} подключения не удалась: {e}")
            if attempt < max_retries:
                time.sleep(retry_delay)
            else:
                logging.error("Не удалось подключиться к Outlook после всех попыток.")
                raise


def debug_outlook_folders(namespace):
    """Вывод структуры всех аккаунтов и папок в Outlook"""
    try:
        for account in namespace.Folders:
            logging.info(f"Учетная запись: {account.Name}")
            for folder in account.Folders:
                logging.info(f"  - Папка: {folder.Name}")
    except Exception as e:
        logging.error(f"Ошибка при выводе папок Outlook: {e}")


def select_calendar_folder(account):
    """Интерактивный выбор папки календаря, если их несколько"""
    calendar_folders = [f for f in account.Folders if "calendar" in f.Name.lower() or "календарь" in f.Name.lower()]

    if len(calendar_folders) == 0:
        logging.warning(f"Учетная запись '{account.Name}' не содержит папки календаря.")
        print("Доступные папки:")
        for i, folder in enumerate(account.Folders, start=1):
            print(f"{i}. {folder.Name}")
        while True:
            try:
                choice = int(input("Выберите папку для добавления событий: "))
                if 1 <= choice <= len(account.Folders):
                    return account.Folders[choice - 1]
                else:
                    print("Неверный номер. Попробуйте снова.")
            except ValueError:
                print("Введите корректный номер.")

    if len(calendar_folders) == 1:
        logging.info(f"Найдена папка календаря: {calendar_folders[0].Name}")
        return calendar_folders[0]

    logging.info(f"Найдено несколько папок календаря в аккаунте '{account.Name}':")
    for i, f in enumerate(calendar_folders, start=1):
        print(f"{i}. {f.Name}")

    while True:
        try:
            choice = int(input("Выберите папку календаря (номер): "))
            if 1 <= choice <= len(calendar_folders):
                selected = calendar_folders[choice - 1]
                logging.info(f"Выбрана папка календаря: {selected.Name}")
                return selected
            else:
                print("Неверный номер. Попробуйте снова.")
        except ValueError:
            print("Введите корректный номер.")


# === Диагностика структуры папок ===
def list_outlook_accounts(namespace):
    """Возвращает список всех аккаунтов и их папок"""
    accounts = []
    for account in namespace.Folders:
        subfolders = [folder.Name for folder in account.Folders]
        accounts.append({
            "account": account.Name,
            "folders": subfolders
        })
    return accounts


def select_account(namespace):
    """Интерактивный выбор аккаунта, если их несколько"""
    accounts = list(namespace.Folders)

    if len(accounts) == 0:
        raise RuntimeError("В Outlook нет учетных записей!")

    if len(accounts) == 1:
        logging.info(f"Найден один аккаунт: {accounts[0].Name}")
        return accounts[0]

    logging.info("Найдено несколько учетных записей в Outlook:")
    for i, acc in enumerate(accounts, start=1):
        print(f"{i}. {acc.Name}")

    while True:
        try:
            choice = int(input("Выберите учетную запись (номер): "))
            if 1 <= choice <= len(accounts):
                selected_account = accounts[choice - 1]
                logging.info(f"Выбран аккаунт: {selected_account.Name}")
                return selected_account
            else:
                print("Неверный номер. Попробуйте еще раз.")
        except ValueError:
            print("Введите корректный номер.")


# === Функция удаления всех событий, созданных скриптом ===
def delete_old_events(calendar, delete_all=False):
    """Удаляет события, созданные скриптом (по категории AutoImportSchedule)."""
    cutoff_date = datetime.now().date() - timedelta(days=60)
    items = calendar.Items
    items.IncludeRecurrences = False
    items.Sort("[Start]")

    # Ограничиваем поиск только событиями начиная с cutoff_date
    restriction = "[Start] >= '" + cutoff_date.strftime("%m/%d/%Y") + "'"
    filtered_items = items.Restrict(restriction)

    deleted_count = 0
    for appt in list(filtered_items):
        try:
            categories = (appt.Categories or "").split(",") if appt.Categories else []
            categories = [c.strip() for c in categories]

            # Удаляем только события с нужной категорией
            if CATEGORY_NAME in categories:
                logging.info(f"Событие на удаление: {appt.Subject} | {appt.Start} | Категории: {appt.Categories}")
                appt.Delete()
                deleted_count += 1
            filtered_items.pop()
        except Exception as e:
            logging.warning(f"Ошибка при удалении события: {e}")

    if delete_all:
        logging.info(f"Удалено ВСЕ события категории {CATEGORY_NAME}: {deleted_count}")
        print(f"Удалено ВСЕ события категории {CATEGORY_NAME}: {deleted_count}")
    else:
        logging.info(f"Удалено {deleted_count} старых событий категории {CATEGORY_NAME}.")
        print(f"Удалено {deleted_count} старых событий категории {CATEGORY_NAME}.")




# === Добавление новых событий ===
def add_events(calendar, records, invitees, offset_hours=3):
    """
    Добавляет события в календарь. Временно применяет смещение offset_hours к Start/End.
    offset_hours: целое число часов, которое прибавляется к времени события (можно 0 чтобы отключить).
    """
    added_count = 0
    outlook = win32.Dispatch("Outlook.Application")

    # Получаем коллекцию TimeZones и попробуем явно взять московскую TZ,
    # если не найдётся — используем CurrentTimeZone как fallback.
    tzs = None
    try:
        tzs = outlook.TimeZones
        try:
            moscow_tz = tzs.Item("Russian Standard Time")
        except Exception:
            try:
                moscow_tz = tzs.Item("Europe/Moscow")
            except Exception:
                moscow_tz = tzs.CurrentTimeZone if hasattr(tzs, "CurrentTimeZone") else None
    except Exception:
        moscow_tz = None

    for event in records:
        # наивный datetime (без tzinfo)
        start_dt = datetime.combine(event["Date"], event["StartTime"])
        end_dt = datetime.combine(event["Date"], event["EndTime"])

        # Применяем временное смещение (workaround)
        if offset_hours:
            orig_start = start_dt
            orig_end = end_dt
            start_dt = start_dt + timedelta(hours=offset_hours)
            end_dt = end_dt + timedelta(hours=offset_hours)
            logging.info(f"Offset {offset_hours}h applied: '{event['Title']}' start {orig_start} -> {start_dt}, end {orig_end} -> {end_dt}")

        # Защита: если после смещения конец <= началу, делаем продолжительность 1 час
        if end_dt <= start_dt:
            logging.warning(f"End <= Start для события '{event['Title']}' после смещения — исправляю (добавляю 1 час).")
            end_dt = start_dt + timedelta(hours=1)

        try:
            appt = calendar.Items.Add()
            appt.Subject = event["Title"]

            # Передаем наивные datetime (Outlook ожидает без tzinfo)
            appt.Start = start_dt
            appt.End = end_dt

            # Пробуем явно указать таймзону, если смогли получить объект
            try:
                if moscow_tz is not None:
                    appt.StartTimeZone = moscow_tz
                    appt.EndTimeZone = moscow_tz
            except Exception as e:
                logging.debug(f"Не удалось установить Start/End TimeZone: {e}")

            appt.ReminderMinutesBeforeStart = 15
            appt.ReminderSet = True
            appt.BusyStatus = 2  # Busy
            appt.Categories = CATEGORY_NAME

            # Добавляем участников
            for email in invitees:
                try:
                    recipient = appt.Recipients.Add(email)
                    if recipient:
                        recipient.Type = 1  # Required attendee
                    else:
                        logging.warning(f"Не удалось добавить участника: {email}")
                    # освобождаем recipient объект явно
                    try:
                        del recipient
                    except Exception:
                        pass
                except Exception as e:
                    logging.warning(f"Ошибка при добавлении участника {email}: {e}")

            # Сохраняем и освобождаем объект
            appt.Save()
            added_count += 1
            logging.info(f"Добавлено событие: {event['Title']} {start_dt} - {end_dt}")

        except Exception as e:
            logging.error(f"Ошибка при добавлении события '{event['Title']}': {e}")
        finally:
            # обязательно освобождаем COM-объект
            try:
                del appt
            except Exception:
                pass

    logging.info(f"Добавлено всего {added_count} событий.")
    print(f"Добавлено {added_count} событий.")
    return added_count


# === Основная функция ===
def main():
    parser = argparse.ArgumentParser(description="Импорт расписания в Outlook.")
    parser.add_argument("--excel-file", help="Путь к Excel-файлу с расписанием")
    parser.add_argument("--invitees-file", help="Файл со списком email для приглашений")
    parser.add_argument("--invitees", help="Список email через запятую")
    parser.add_argument("--delete-all", action="store_true", help="Удалить все события, созданные скриптом, без добавления новых")
    parser.add_argument("--offset-hours", type=int, default=3,
                        help="Временный workaround: прибавлять N часов к Start/End (по умолчанию 3). Установите 0 чтобы отключить.")

    args = parser.parse_args()

    # Путь к скрипту или exe
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

    # Excel по умолчанию
    excel_file = args.excel_file or os.path.join(base_dir, DEFAULT_EXCEL_NAME)
    if not os.path.exists(excel_file):
        print(f"Файл Excel не найден: {excel_file}")
        sys.exit(1)

    # Лог-файл
    log_file = os.path.join(base_dir, "script_log.txt")
    setup_logger(log_file)

    logging.info("===== Запуск скрипта =====")

    # Загружаем участников
    invitees = load_invitees(args.invitees_file, args.invitees)
    logging.info(f"Загружено {len(invitees)} участников: {', '.join(invitees) if invitees else 'нет'}")

    # Подключение к Outlook
    outlook, namespace, calendar = connect_to_outlook()

    # Удаление всех или старых событий
    if args.delete_all:
        delete_old_events(calendar, delete_all=True)
        logging.info("Режим удаления завершён.")
        return

    # Загружаем расписание
    df = load_excel(excel_file)
    records = parse_schedule(df)
    logging.info(f"Найдено {len(records)} событий в расписании.")

    # Сначала удаляем старые
    delete_old_events(calendar, delete_all=False)

    # Добавляем новые события
    add_events(calendar, records, invitees)


if __name__ == "__main__":
    main()
