import os
import re
import sys
import time
from typing import Dict, List, Any

# Библиотеки автоматизации и работы с документами
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Настройка вывода UTF-8 для корректного отображения кириллицы в консоли Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass


def init_browser() -> uc.Chrome:
    """Инициализация браузера с защитой от детекта автоматизации."""
    print(" [1/7] Запуск браузера Chrome (undetected-chromedriver)...")
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    
    driver = uc.Chrome(options=options)
    return driver


def close_overlays(driver: uc.Chrome):
    """Закрытие всплывающих окон города и куки, если они появились."""
    time.sleep(1)
    # Подтверждение города (кнопка 'Все верно')
    try:
        city_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Все верно')]")
        if city_buttons:
            city_buttons[0].click()
            time.sleep(0.5)
    except Exception:
        pass

    # Предупреждение о cookie (кнопка 'Понятно')
    try:
        cookie_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Понятно')]")
        if cookie_buttons:
            cookie_buttons[0].click()
            time.sleep(0.5)
    except Exception:
        pass

    # Окно с информацией об избранном
    try:
        close_modal = driver.find_elements(By.XPATH, "//button[contains(@class, 'modal__close') or @aria-label='Закрыть']")
        if close_modal:
            close_modal[0].click()
            time.sleep(0.5)
    except Exception:
        pass


def search_products(driver: uc.Chrome, search_term: str = "мышь"):
    """Шаг 1: Поиск товаров через поисковую строку."""
    print(f" [2/7] Открытие сайта https://www.dns-shop.ru/ и поиск по запросу '{search_term}'...")
    driver.get("https://www.dns-shop.ru/")
    
    wait = WebDriverWait(driver, 20)
    close_overlays(driver)

    # Поиск поля ввода
    search_input = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//input[@type='search' or contains(@placeholder, 'Поиск')]"))
    )
    search_input.click()
    time.sleep(0.5)
    search_input.clear()
    search_input.send_keys(search_term)
    time.sleep(0.5)
    search_input.send_keys(Keys.ENTER)

    # Ожидание загрузки каталога товаров
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-id='product'], .catalog-product")))
    close_overlays(driver)
    time.sleep(2)
    print("   -> Результаты поиска успешно загружены.")


def sort_by_reviews(driver: uc.Chrome):
    """Шаг 2: Изменение сортировки на 'По количеству отзывов'."""
    print(" [3/7] Изменение сортировки на 'По количеству отзывов'...")
    wait = WebDriverWait(driver, 15)
    close_overlays(driver)

    # Клик по блоку сортировки
    sort_dropdown = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//*[contains(@class, 'top-filter')]//span[contains(text(), 'Сортировка')]/.. | //*[contains(@class, 'top-filter__label')]/.."))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sort_dropdown)
    time.sleep(0.5)
    sort_dropdown.click()
    time.sleep(1)

    # Выбор пункта 'По количеству отзывов'
    review_option = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'По количеству отзывов')] | //span[contains(., 'По количеству отзывов')]"))
    )
    review_option.click()
    print("   -> Сортировка переключена, ожидание обновления каталога...")
    time.sleep(4)
    close_overlays(driver)


def parse_page_products(driver: uc.Chrome) -> List[Dict[str, Any]]:
    """Извлечение товаров с первой страницы каталога."""
    print(" [4/7] Сбор информации о товарах с первой страницы...")
    product_cards = driver.find_elements(By.CSS_SELECTOR, "[data-id='product'], .catalog-product")
    print(f"   -> Найдено карточек товаров: {len(product_cards)}")

    products: List[Dict[str, Any]] = []

    for index, card in enumerate(product_cards):
        try:
            # Название товара
            title_el = card.find_element(By.CSS_SELECTOR, "a.catalog-product__name, [data-role='product-title'], a[class*='name']")
            title = title_el.text.strip()
            if not title:
                continue

            # Цена товара
            price = 0
            try:
                price_el = card.find_element(By.CSS_SELECTOR, ".product-buy__price, [class*='price__current'], [class*='buy__price']")
                price_text = price_el.text.strip()
                match_price = re.search(r"(\d[\d\s]*)\s*₽", price_text)
                if match_price:
                    price = int(re.sub(r"\s+", "", match_price.group(1)))
                else:
                    digits = re.sub(r"[^\d]", "", price_text)
                    price = int(digits) if digits else 0
            except Exception:
                price = 0

            # Рейтинг товара
            rating = 0.0
            try:
                rating_el = card.find_element(By.CSS_SELECTOR, "a.catalog-product__rating, [data-rating], [class*='rating']")
                rating_text = rating_el.text.strip()
                match_rating = re.search(r"(\d+(?:[.,]\d+)?)", rating_text)
                if match_rating:
                    rating = float(match_rating.group(1).replace(",", "."))
            except Exception:
                rating = 0.0

            # Кнопка добавления в избранное
            fav_btn = None
            try:
                btns = card.find_elements(By.CSS_SELECTOR, "button.button-ui_white, button[class*='wishlist'], button[class*='favorite'], button[class*='like']")
                for b in btns:
                    label = (b.get_attribute("title") or b.get_attribute("aria-label") or "").lower()
                    cls = (b.get_attribute("class") or "").lower()
                    if "избранн" in label or "like" in cls or "wishlist" in cls:
                        fav_btn = b
                        break
                if not fav_btn and btns:
                    fav_btn = btns[0]
            except Exception:
                pass

            products.append({
                "index": index,
                "title": title,
                "price": price,
                "rating": rating,
                "fav_btn": fav_btn,
                "status": "-"
            })
        except Exception as err:
            print(f"   [Предупреждение] Ошибка парсинга карточки {index}: {err}")

    print(f"   -> Успешно обработано товаров: {len(products)}")
    return products


def add_cheapest_to_favorites(driver: uc.Chrome, products: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Шаг 3: Поиск самого дешевого товара и добавление его в избранное."""
    print(" [5/7] Поиск самого дешевого товара и добавление в избранное...")
    valid_products = [p for p in products if p["price"] > 0]
    if not valid_products:
        raise RuntimeError("Не удалось найти товары с указанной ценой!")

    cheapest = min(valid_products, key=lambda x: x["price"])
    cheapest["status"] = "+"
    print(f"   -> Самый дешевый товар: '{cheapest['title']}' | Цена: {cheapest['price']} руб.")

    if cheapest["fav_btn"]:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cheapest["fav_btn"])
        time.sleep(1)
        try:
            cheapest["fav_btn"].click()
        except Exception:
            driver.execute_script("arguments[0].click();", cheapest["fav_btn"])
        time.sleep(2)
        print("   -> Товар успешно добавлен в избранное!")
        close_overlays(driver)
    else:
        print("   [Внимание] Кнопка добавления в избранное для данного товара не найдена.")

    return cheapest


def export_to_excel(products: List[Dict[str, Any]], filename: str = "products.xlsx"):
    """
    Часть 2: Формирование документа .xlsx со всеми требованиями:
    1) Формат: Название, Цена, Рейтинг, Статус.
    2) Заливка шапки цветом.
    3) Сортировка по названию в алфавитном порядке + фильтр Excel.
    4) Формула MIN по столбцу 'Цена'.
    5) Условное форматирование 'Рейтинг': >= 4.8 зеленый, < 4.8 красный.
    6) Условное форматирование 'Статус': '+' зеленый, '-' красный + выпадающий список.
    """
    print(f" [6/7] Формирование отчета Excel ({filename})...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Товары"

    # Требование 3: Таблица отсортирована по названию в алфавитном порядке
    sorted_products = sorted(products, key=lambda x: x["title"].lower())

    # Требование 1: Шапка таблицы
    headers = ["Название", "Цена", "Рейтинг", "Статус"]
    ws.append(headers)

    # Добавление строк данных
    for p in sorted_products:
        ws.append([p["title"], p["price"], p["rating"], p["status"]])

    last_row = len(sorted_products) + 1  # последняя строка данных (с учетом шапки)

    # Требование 2: Шапку таблицы нужно залить цветом
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")  # Синий корпоративный
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Форматирование ячеек данных
    for row in range(2, last_row + 1):
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=row, column=2).number_format = '#,##0 "₽"'
        ws.cell(row=row, column=3).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row, column=3).number_format = '0.00'
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="center", vertical="center")

        for col in range(1, 5):
            ws.cell(row=row, column=col).border = thin_border

    # Требование 3: Фильтр Excel со связкой сортировки
    ws.auto_filter.ref = f"A1:D{last_row}"
    ws.auto_filter.add_sort_condition(f"A2:A{last_row}", descending=False)

    # Требование 4: После последней ячейки в столбце 'Цена' - формула MIN
    min_row = last_row + 1
    cell_min_label = ws.cell(row=min_row, column=1, value="Минимальная цена")
    cell_min_label.font = Font(name="Calibri", size=11, bold=True)
    cell_min_label.alignment = Alignment(horizontal="right", vertical="center")

    cell_min_value = ws.cell(row=min_row, column=2, value=f"=MIN(B2:B{last_row})")
    cell_min_value.font = Font(name="Calibri", size=11, bold=True)
    cell_min_value.number_format = '#,##0 "₽"'
    cell_min_value.alignment = Alignment(horizontal="right", vertical="center")

    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    cell_min_label.fill = total_fill
    cell_min_value.fill = total_fill
    cell_min_label.border = thin_border
    cell_min_value.border = thin_border

    # Цветовые стили для условного форматирования
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100", bold=True)
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)

    # Требование 5: Условное форматирование столбца 'Рейтинг' (C2:C{last_row})
    rule_rating_green = CellIsRule(operator="greaterThanOrEqual", formula=["4.8"], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_rating_red = CellIsRule(operator="lessThan", formula=["4.8"], stopIfTrue=True, fill=red_fill, font=red_font)
    ws.conditional_formatting.add(f"C2:C{last_row}", rule_rating_green)
    ws.conditional_formatting.add(f"C2:C{last_row}", rule_rating_red)

    # Требование 6: Условное форматирование столбца 'Статус' (D2:D{last_row})
    rule_status_plus = CellIsRule(operator="equal", formula=['"+"'], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_status_minus = CellIsRule(operator="equal", formula=['"-"'], stopIfTrue=True, fill=red_fill, font=red_font)
    ws.conditional_formatting.add(f"D2:D{last_row}", rule_status_plus)
    ws.conditional_formatting.add(f"D2:D{last_row}", rule_status_minus)

    # Требование 6: Выпадающий список выбора только из '+' и '-'
    dv = DataValidation(type="list", formula1='"+,-"', allow_blank=False)
    dv.errorTitle = "Недопустимый статус"
    dv.error = "Выберите значение из выпадающего списка: '+' или '-'"
    dv.promptTitle = "Статус товара"
    dv.prompt = "Выберите '+' (в избранном) или '-' (не в избранном)"
    ws.add_data_validation(dv)
    dv.add(f"D2:D{last_row}")

    # Автоматическая настройка ширины столбцов
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or "")
            if cell.number_format and "₽" in cell.number_format and isinstance(cell.value, (int, float)):
                val_str = f"{cell.value:,} ₽"
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Комфортная ширина для столбца с названием
    ws.column_dimensions["A"].width = 50

    saved = False
    current_filename = filename
    while not saved:
        try:
            wb.save(current_filename)
            print(f"   -> Файл '{current_filename}' успешно сохранен со всеми требованиями.")
            saved = True
        except PermissionError:
            fallback_filename = f"products_{int(time.time())}.xlsx"
            print(f"\n   [Внимание!] Файл '{current_filename}' сейчас открыт в Excel (или другой программе).")
            print("   Windows блокирует перезапись открытых файлов.")
            print(f"   Пожалуйста, закройте '{current_filename}' в Excel и нажмите Enter для повторной попытки,")
            print(f"   или введите 'new', чтобы сохранить файл как '{fallback_filename}': ", end="", flush=True)
            try:
                choice = input().strip().lower()
                if choice in ("new", "n", "yes", "y", "нов", "новый"):
                    current_filename = fallback_filename
            except Exception:
                current_filename = fallback_filename


def go_to_wishlist(driver: uc.Chrome):
    """Шаг 4: Переход на вкладку 'Избранное' и удержание окна открытым."""
    print(" [7/7] Переход на вкладку 'Избранное'...")
    try:
        fav_link = driver.find_element(By.XPATH, "//a[contains(@href, 'wishlist') or contains(., 'Избранное')]")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", fav_link)
        time.sleep(0.5)
        fav_link.click()
    except Exception:
        driver.get("https://www.dns-shop.ru/profile/wishlist/")

    time.sleep(3)
    close_overlays(driver)
    print("\n" + "="*70)
    print(" РОБОТ УСПЕШНО ВЫПОЛНИЛ ВСЕ ШАГИ ЗАДАНИЯ!")
    print(f" Браузер открыт на странице: {driver.current_url}")
    print(" Согласно требованию шага 4, окно браузера НЕ закрывается.")
    print(" Чтобы завершить работу программы, нажмите Enter в этой консоли...")
    print("="*70 + "\n")


def main():
    search_query = "мышь"
    if len(sys.argv) > 1:
        search_query = " ".join(sys.argv[1:])

    driver = None
    try:
        # Часть 1: Шаги 1-3
        driver = init_browser()
        search_products(driver, search_query)
        sort_by_reviews(driver)
        products = parse_page_products(driver)
        
        if not products:
            print("Ошибка: не удалось найти товары на странице!")
            return

        cheapest = add_cheapest_to_favorites(driver, products)

        # Часть 2: Выполняется после шага 3 части 1
        export_to_excel(products, "products.xlsx")

        # Часть 1: Шаг 4 - переход в избранное без закрытия окна
        go_to_wishlist(driver)

        # Ожидание действия пользователя для остановки без закрытия
        try:
            input("Нажмите Enter для завершения программы и закрытия браузера...")
        except (EOFError, KeyboardInterrupt):
            pass

    except Exception as e:
        print(f"\n[Критическая ошибка]: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
            driver.__del__ = lambda: None
            print("Браузер закрыт.")


if __name__ == "__main__":
    main()
