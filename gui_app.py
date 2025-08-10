from pywinauto.keyboard import send_keys
import datetime
import os
import sys
import io
import locale

import subprocess
import time
import psutil
import pyautogui
from pywinauto import Application, timings, Desktop, mouse
from pywinauto.findwindows import ElementNotFoundError
from const import IMAGE_PATH, APP_TITLE, APP_SIZE, LOG_FILE
from excel_parser import parse_excel_schedule, TrainScheduleEntry

def arm_window_connect(self):
    arm_app = None
    arm_window = None

    # Способ 1: Поиск среди всех окон
    try:
        desktop = Desktop(backend="uia")
        for window in desktop.windows():
            title = window.window_text()
            if "АРМ Нарядчика" in title:
                self.log(f"✅ Найдено окно АРМ Нарядчика: {title}")
                # Подключаемся к процессу этого окна
                window_pid = window.process_id()
                arm_app = Application(backend="uia").connect(process=window_pid)
                arm_window = arm_app.window(title_re=".*АРМ Нарядчика.*")
                return arm_window
    except Exception as search_err:
        self.log(f"⚠️ Ошибка поиска окна: {search_err}")

    # Способ 2: Прямое подключение (если не нашли выше)
    if not arm_app:
        try:
            arm_app = Application(backend="uia").connect(title_re=".*АРМ Нарядчика.*", timeout=10)
            arm_window = arm_app.window(title_re=".*АРМ Нарядчика.*")
            self.log("✅ Подключение к АРМ Нарядчика выполнено")
            return arm_window
        except Exception as connect_err:
            self.log(f"❌ Не удалось подключиться к АРМ Нарядчика: {connect_err}")
            return None

    if not arm_window:
        self.log("❌ Окно АРМ Нарядчика не найдено")
        return None

    self.log("✅ Окно АРМ Нарядчика найдено и готово")
    arm_window.wait('ready', timeout=10)
    self.log("✅ Главное окно АРМ Нарядчика готово к работе")
    return arm_window

def start_client(self):
    """Автоматизация работы с ClientManager и АРМ Нарядчика."""
    try:
        # 1. Запуск Client Manager
        self.log("🔄 Запуск Client Manager...")
        app = Application(backend="uia").start(r"C:\ClientManager\ClientsCache\ClientManager\ClientManager.exe")
        time.sleep(5)
        
        # 2. Работа с окном соединения
        clientmanager = app.window(title_re=".*Соединение.*")
        if clientmanager.exists():
            self.log('✅ Обнаружил окно приложения "Соединение"')
        else:
            self.log('❌ Не обнаружил окно "Соединение"')
            return
            
        apply = clientmanager.child_window(auto_id="btnOk")
        if apply.exists():
            self.log('✅ Обнаружил кнопку "Применить"')
            apply.click()
            self.log('✅ Кнопка "Применить" нажата')
        else:
            self.log('❌ Не обнаружил кнопку "Применить"')
            return
            
        time.sleep(2)
        
        # 3. Работа с окном "Управление клиентом"
        window_client = app.window(title_re=".*Управление.*")
        time.sleep(2)
        if window_client.exists():
            self.log('✅ Обнаружил окно "Управление клиентом"')
        else:
            self.log('❌ Не обнаружил окно "Управление клиентом"')
            return

        # 4. Поиск и запуск АРМ Нарядчика
        try:
            # Подключение к приложению "Управление клиентом"
            app_client = Application(backend="uia").connect(title_re=".*Управление клиентом.*", timeout=20)
            main_window = app_client.window(title_re=".*Управление клиентом.*")
            main_window.wait('visible', timeout=7)
            
            # Поиск списка
            list_view = main_window.child_window(
                class_name="WindowsForms10.SysListView32.app.0.378734a",
                control_type="List"
            )
            
            if not list_view.exists():
                self.log("❌ Список не найден")
                return

            # Поиск элемента "АРМ Нарядчика"
            target_item = None
            items = list_view.descendants(control_type="ListItem")
            
            for item in items:
                item_text = item.window_text()
                self.log(f"Найден элемент списка: {item_text}")
                if "Нарядчика" in item_text:
                    target_item = item
                    self.log(f'✅ Найден элемент "АРМ Нарядчика": {item_text}')
                    break
            
            if target_item:
                # Попробуем разные методы клика
                try:
                    # Метод 1: Двойной клик
                    target_item.click_input(double=True)
                    self.log("✅ Двойной клик по элементу выполнен")
                except Exception as click_err:
                    self.log(f"⚠️ Ошибка двойного клика: {click_err}")
                    try:
                        # Метод 2: Клик правой кнопкой + контекстное меню
                        target_item.click_input(button='right')
                        time.sleep(0.5)
                        # Ищем пункт "Запустить" в контекстном меню
                        context_menu = app_client.window(title_re=".*")
                        run_item = context_menu.child_window(title="Запустить", control_type="MenuItem")
                        if run_item.exists():
                            run_item.click_input()
                            self.log("✅ Запуск через контекстное меню")
                        else:
                            self.log("❌ Пункт 'Запустить' в контекстном меню не найден")
                    except Exception as context_err:
                        self.log(f"⚠️ Ошибка контекстного меню: {context_err}")
                        # Метод 3: Клик по координатам (если знаем примерные координаты)
                        try:
                            from pywinauto import mouse
                            rect = target_item.rectangle()
                            center_x = rect.left + (rect.width() // 2)
                            center_y = rect.top + (rect.height() // 2)
                            mouse.double_click(coords=(center_x, center_y))
                            self.log("✅ Двойной клик по координатам выполнен")
                        except Exception as coord_err:
                            self.log(f"❌ Все методы клика не сработали: {coord_err}")
                            return
            else:
                self.log("❌ Элемент 'АРМ Нарядчика' не найден в списке")
                return

        except Exception as ex:
            self.log(f"❌ Ошибка при работе с окном 'Управление клиентом': {str(ex)}")
            return

        # 5. Ожидание запуска АРМ Нарядчика
        self.log("⏳ Ожидание запуска АРМ Нарядчика...")
        time.sleep(5)

        # 6. Работа с окном "АРМ Нарядчика"
        self.log("🔍 Поиск окна 'АРМ Нарядчика'...")
        
        arm_window = arm_window_connect(self)
        if not arm_window:
            return

        # 7. Работа с меню "Расписания" - ИСПРАВЛЕННАЯ ЧАСТЬ
        self.log("🔍 Поиск пункта меню 'Расписания'...")
        
        # Ищем все элементы MenuBar и работаем с первым подходящим
        try:
            menu_bars = arm_window.children(control_type="MenuBar")
            self.log(f"✅ Найдено {len(menu_bars)} строк(и) меню")
            
            schedules_menu = None
            menu_strip = None
            
            # Перебираем все найденные MenuBar
            for i, mb in enumerate(menu_bars):
                self.log(f"🔍 Проверка MenuBar #{i+1}")
                try:
                    # Ищем пункт "Расписания" в этой строке меню
                    potential_schedules = mb.child_window(title="Расписания", control_type="MenuItem")
                    if potential_schedules.exists():
                        schedules_menu = potential_schedules
                        menu_strip = mb
                        self.log(f"✅ Пункт 'Расписания' найден в MenuBar #{i+1}")
                        break
                except Exception as mb_err:
                    self.log(f"⚠️ Ошибка при проверке MenuBar #{i+1}: {mb_err}")
                    continue
            
            # Если не нашли через child_window, ищем напрямую в главном окне
            if not schedules_menu:
                self.log("⚠️ Пункт 'Расписания' не найден в MenuBar, ищу напрямую...")
                schedules_menu = arm_window.child_window(title="Расписания", control_type="MenuItem")
                
        except Exception as menu_search_err:
            self.log(f"⚠️ Ошибка поиска MenuBar: {menu_search_err}")
            # Последний способ - ищем напрямую
            schedules_menu = arm_window.child_window(title="Расписания", control_type="MenuItem")

        if schedules_menu and schedules_menu.exists():
            self.log("✅ Пункт меню 'Расписания' найден.")
            self.log("🖱️ Нажатие на 'Расписания'...")
            schedules_menu.click_input()
            self.log("✅ Клик по 'Расписания' выполнен.")
        else:
            self.log("❌ Пункт меню 'Расписания' НЕ НАЙДЕН!")
            # Диагностика - показываем все найденные MenuItem
            try:
                all_menu_items = arm_window.descendants(control_type="MenuItem")
                self.log("🔍 Все найденные пункты меню:")
                for i, item in enumerate(all_menu_items):
                    item_text = item.window_text()
                    if item_text.strip():  # Показываем только непустые элементы
                        self.log(f"  {i+1}. '{item_text}'")
            except Exception as diag_err:
                self.log(f"❌ Ошибка диагностики меню: {diag_err}")
            return

        # 8. Работа с подменю "График обслуживания поездов МВПС"
        self.log("🔍 Поиск подменю 'График обслуживания поездов МВПС'...")
        mvs_schedule_item = arm_window.child_window(title="График обслуживания поездов МВПС", control_type="MenuItem")

        # Ждём появления элемента (на случай задержки отображения меню)
        from pywinauto.timings import wait_until_passes
        try:
            wait_until_passes(7, 0.5, lambda: mvs_schedule_item.exists())
            self.log("✅ Подменю 'График обслуживания поездов МВПС' найдено.")
            self.log("🖱️ Нажатие на 'График обслуживания поездов МВПС'...")
            mvs_schedule_item.click_input()
            self.log("✅ Клик по 'График обслуживания поездов МВПС' выполнен.")
            self.log("🎉 Все операции успешно выполнены!")
        except Exception as e:
            self.log(f"❌ Подменю 'График обслуживания поездов МВПС' НЕ НАЙДЕНО или не открылось: {str(e)}")
            # Попробуем найти все подменю для диагностики
            try:
                schedules_menu.click_input()  # Открываем меню снова
                time.sleep(1)
                submenu_items = arm_window.descendants(control_type="MenuItem")
                self.log("🔍 Все доступные пункты меню после открытия 'Расписания':")
                for item in submenu_items:
                    item_text = item.window_text()
                    if item_text.strip() and "Расписания" not in item_text:  # Исключаем сам пункт "Расписания"
                        self.log(f"  - '{item_text}'")
            except Exception as submenu_err:
                self.log(f"❌ Ошибка при поиске подменю: {submenu_err}")

    except Exception as e:
        self.log(f"❌ Критическая ошибка в скрипте: {str(e)}")
        import traceback
        self.log(f"Трассировка ошибки: {traceback.format_exc()}")

def open_image(app_instance):
    """Открытие изображения с улучшенным поиском окна"""
    try:
        if not os.path.exists(app_instance.IMAGE_PATH):
            app_instance.log(f"❌ Файл не найден: {app_instance.IMAGE_PATH}")
            return

        app_instance.log("🖱️ Открытие изображения...")
        subprocess.Popen(f'explorer "{app_instance.IMAGE_PATH}"', shell=True)
        time.sleep(3)  # Даем время для открытия приложения

        try:
            # Пробуем несколько вариантов названий окон
            app = None
            window_titles = [
                ".*Фотографии.*",
                ".*Photos.*",
                ".*Просмотр фотографий.*",
                ".*Viewer.*"
            ]

            for title in window_titles:
                try:
                    app = Application(backend="uia").connect(title_re=title)
                    break
                except ElementNotFoundError:
                    continue

            if app:
                app_instance.log("✅ Изображение открыто в просмотрщике")
                # Можно добавить работу с элементами окна здесь
            else:
                app_instance.log("⚠️ Не удалось подключиться к окну просмотрщика")
                app_instance.log("ℹ️ Проверьте, открылось ли изображение вручную")

        except Exception as e:
            app_instance.log(f"⚠️ Ошибка подключения к приложению: {str(e)}")

    except Exception as e:
        app_instance.log(f"❌ Критическая ошибка: {str(e)}")

def get_toolbar_button_by_index(arm_window, physical_index_in_toolbar):
    """
    Получает кнопку из toolStrip1 по её физическому порядковому индексу в тулбаре.
    Ищет toolStrip1 внутри окна FrmMVPSTimetable.
    Использует упрощённый и надёжный способ.

    :param arm_window: Главное окно приложения АРМ Нарядчика.
    :param physical_index_in_toolbar: Физический индекс кнопки в toolStrip1 (начиная с 0).
    :return: Объект кнопки или None, если не найдена.
    """
    try:
        print(f"--- Начало поиска кнопки с физическим индексом {physical_index_in_toolbar} ---")

        # 1. Находим внутреннее окно "График обслуживания поездов МВПС" по его AutomationId.
        frm_mvsp_timetable_spec = arm_window.child_window(
            auto_id="FrmMVPSTimetable",  # Используем auto_id
            control_type="Window",
            found_index=0
        )
        frm_mvsp_timetable_spec.wait('exists', timeout=7)
        print("✅ Спецификация внутреннего окна FrmMVPSTimetable получена и существует.")
        frm_mvsp_timetable = frm_mvsp_timetable_spec.wrapper_object()
        print("✅ Объект внутреннего окна FrmMVPSTimetable получен.")

        # 2. Внутри этого окна находим toolStrip1.
        # Используем auto_id для надёжности.
        toolstrip_spec = frm_mvsp_timetable_spec.child_window(
            auto_id="ts1",  # Используем auto_id для toolStrip1
            control_type="ToolBar",
            found_index=0
        )
        toolstrip_spec.wait('exists', timeout=7)
        toolstrip = toolstrip_spec.wrapper_object()
        print("✅ ToolStrip1 найден внутри FrmMVPSTimetable по AutoId.")

        # 3. Получаем список ВСЕХ потомков toolStrip1.
        # Это список всех элементов в порядке их следования в тулбаре.
        all_toolbar_children = toolstrip.children()
        print(f"   В toolStrip1 найдено {len(all_toolbar_children)} потомков.")

        # 4. Проверяем, что запрашиваемый физический индекс существует.
        if 0 <= physical_index_in_toolbar < len(all_toolbar_children):
            button = all_toolbar_children[physical_index_in_toolbar]
            # Для диагностики, выведем информацию о найденном элементе
            try:
                btn_title = button.window_text()
                btn_ctrl_type = button.element_info.control_type
                print(f"✅ Элемент с физическим индексом {physical_index_in_toolbar} найден.")
                print(f"   Название элемента: '{btn_title}'")
                print(f"   Тип элемента: '{btn_ctrl_type}'")
                # Проверим, является ли он кнопкой. Это для информации, не для фильтрации.
                # Если он не кнопка, это будет видно по названию и типу.
            except Exception as info_err:
                print(f"   Ошибка получения информации об элементе: {info_err}")
            print("--- Конец поиска кнопки ---")
            return button
        else:
            print(
                f"❌ Физический индекс {physical_index_in_toolbar} выходит за пределы списка потомков (0-{len(all_toolbar_children) - 1}).")
            # Для диагностики выведем все найденные элементы
            for i, child in enumerate(all_toolbar_children):
                try:
                    child_title = child.window_text()
                    child_ctrl_type = child.element_info.control_type
                    print(f"      Элемент #{i}: '{child_title}' (Тип: {child_ctrl_type})")
                except Exception as child_err:
                    print(f"      Элемент #{i}: Ошибка получения информации ({child_err})")
            print("--- Конец поиска кнопки ---")
            return None

    except Exception as e:
        print(f"❌ Ошибка при получении элемента с физическим индексом {physical_index_in_toolbar}: {e}")
        import traceback
        print(traceback.format_exc())
        print("--- Конец поиска кнопки ---")
        return None


def find_exact_row(self, arm_window, expected_route_name="Пробный МЦД-1"):
    """
    Ищет строку в таблице grMain внутри окна "График обслуживания поездов МВПС".
    :param arm_window: Главное окно АРМ Нарядчика.
    :param expected_route_name: Название маршрута для поиска.
    :return: Найденный элемент строки или None.
    """
    try:
        print("--- Начало поиска строки ---")
        # --- Шаг 1: Найти окно FrmMVPSTimetable ---
        print("1. Поиск окна 'График обслуживания поездов МВПС'...")
        frm_mvsp_timetable = arm_window.child_window(
            title="График обслуживания поездов МВПС",
            control_type="Window"
        ).wait('exists', timeout=15)
        print("   ✅ Окно 'График обслуживания поездов МВПС' найдено.")

        # --- Шаг 2: Найти таблицу grMain по AutomationId ---
        print("2. Поиск таблицы 'grMain'...")
        data_grid = None
        # Ищем все таблицы внутри окна
        all_tables = frm_mvsp_timetable.descendants(control_type="Table")
        print(f"   Найдено таблиц: {len(all_tables)}")

        # Фильтруем вручную по AutomationId
        for i, table in enumerate(all_tables):
            try:
                # ВАЖНО: automation_id() это МЕТОД, его нужно вызывать!
                table_auto_id = table.automation_id()
                print(f"   Таблица {i}: AutomationId = '{table_auto_id}'")
                if table_auto_id == "grMain":
                    data_grid = table
                    print("   ✅ Таблица 'grMain' найдена по AutomationId!")
                    break
            except Exception as table_check_err:
                print(f"   ⚠️ Ошибка проверки таблицы {i}: {table_check_err}")
                continue

        if not data_grid:
            print("   ❌ ОШИБКА: Таблица 'grMain' не найдена.")
            # Диагностика: Покажем все найденные элементы
            # frm_mvsp_timetable.print_control_identifiers() # Раскомментируйте для отладки
            return None

        # --- Шаг 3: Найти строку ---
        print("3. Поиск строк в таблице...")
        # Ищем ВСЕХ прямых потомков таблицы, чтобы избежать KeyError
        # Потом отфильтруем по типу вручную, если нужно.
        all_children = data_grid.children()
        print(f"   Найдено прямых потомков таблицы: {len(all_children)}")

        # Фильтруем потомков, оставляя только те, которые могут быть строками.
        # Мы знаем, что строки могут быть DataItem или CustomControl.
        # Так как прямое указание этих типов в children() может вызвать KeyError,
        # мы фильтруем уже полученный список.
        potential_rows = []
        for child in all_children:
            try:
                # Получаем тип элемента из element_info
                ctrl_type = child.element_info.control_type
                # Печатаем для отладки (можно убрать позже)
                # print(f"     Потомок: {child.window_text()}, Тип: {ctrl_type}")
                # Фильтруем по известным типам строк. 'DataItem' и 'CustomControl'
                # это логические имена, реальные имена типов могут отличаться.
                # Проверим на совпадение с известными строковыми типами.
                # UIA_DataItemControlTypeId = 0xC364 (50020)
                # UIA_CustomControlTypeId = 0xC367 (50023) - менее вероятен для строки
                # Но лучше проверить по имени типа.
                if "DataItem" in ctrl_type or "Custom" in ctrl_type:  # Более гибкая проверка
                    potential_rows.append(child)
            except Exception as filter_err:
                # Если не удалось получить тип, пропускаем элемент
                print(f"     ⚠️ Ошибка фильтрации потомка: {filter_err}")
                continue

        # Альтернатива: если фильтрация по типу не работает или не надежна,
        # можно просто взять всех потомков и проверять их свойства Legacy.Value.
        # potential_rows = all_children

        print(f"   Найдено потенциальных строк (после фильтрации): {len(potential_rows)}")

        if not potential_rows:
            print("   ⚠️ В таблице нет потенциальных строк (после фильтрации).")
            # Попробуем без фильтрации
            potential_rows = all_children
            print(f"   Попробуем все {len(potential_rows)} потомков как строки.")

        # --- Шаг 4: Проверить строки на соответствие критерию ---
        print(f"4. Поиск строки с Legacy.Value содержащим '{expected_route_name}'...")
        for i, row in enumerate(potential_rows):
            try:
                # Получаем Legacy.Value
                legacy_value = ""
                try:
                    # Пробуем получить через legacy_properties
                    legacy_props = row.legacy_properties()
                    legacy_value = legacy_props.get('Value', '')
                except Exception as get_legacy_err:
                    print(f"     ⚠️ Ошибка получения Legacy.Value у строки {i}: {get_legacy_err}")
                    continue  # Если не можем получить значение, пропускаем строку

                # print(f"   Проверка строки {i}: Legacy.Value = '{legacy_value}'") # Для отладки

                # Проверяем, содержит ли значение искомое название маршрута
                # Используем str() на случай, если legacy_value не строка
                if expected_route_name in str(legacy_value):
                    self.log(f"     ✅ НАЙДЕНА СТРОКА {i}, содержащая '{expected_route_name}'!")
                    print(
                        f"     Информация о найденной строке: Имя='{row.window_text()}', Тип='{row.element_info.control_type}'")
                    return row

            except Exception as row_error:
                print(f"   ⚠️ Ошибка проверки строки {i}: {row_error}")
                continue

        print(f"   ℹ️ Строка с Legacy.Value содержащим '{expected_route_name}' не найдена.")
        return None

    except Exception as e:
        print(f"   ❌ Критическая ошибка в процессе поиска строки: {e}")
        self.log('Ошибка поиска записи графика')
        import traceback
        traceback.print_exc()
        return None
    finally:
        print("--- Конец поиска строки ---")


def get_input_field(self, arm_window, expected_route_name="Пробный МЦД-1"):
    try:
        # 1. Получаем toolStrip2
        tool_strip2 = arm_window.child_window(
            title="toolStrip2",
            control_type="ToolBar"
        ).wait('exists', timeout=10)

        # 2. Ищем все поля ввода в тулбаре
        edit_controls = tool_strip2.descendants(
            control_type="Edit",
            class_name="WindowsForms10.EDIT.app.0.378734a"
        )

        # 3. Фильтруем по комбинации признаков
        for edit in edit_controls:
            try:
                rect = edit.rectangle()

                # Проверяем координаты (подстройте под ваше приложение)
                coord_ok = (100 <= rect.left <= 200 and 100 <= rect.top <= 150)

                # Проверяем соседний элемент
                next_ctrl = edit.next_sibling()
                next_ok = next_ctrl and "DateTime" in next_ctrl.class_name()

                # Проверяем предыдущий элемент
                prev_ctrl = edit.previous_sibling()
                prev_ok = prev_ctrl and "Label" in prev_ctrl.class_name()

                if coord_ok and (next_ok or prev_ok):
                    edit.draw_outline(colour='green', thickness=2)
                    return edit

            except Exception as e:
                continue

        # 4. Альтернативный поиск по тексту подсказки
        for edit in edit_controls:
            try:
                if "название" in edit.legacy_properties().get('HelpText', '').lower():
                    return edit
            except:
                continue

        # 5. Поиск по порядку в иерархии (если поле всегда первое/второе)
        if len(edit_controls) >= 1:
            return edit_controls[0]  # или 1 для второго поля

    except Exception as e:
        print(f"Ошибка при поиске поля: {e}")

    # 6. Последний вариант - клик по координатам
    try:
        coords = (150, 120)  # Подстройте координаты
        mouse.click(coords=coords)
        time.sleep(0.5)

        # Получаем элемент с фокусом
        focused = Application(backend="uia").connect(active=True).window()
        if "EDIT" in focused.class_name():
            return focused
    except Exception as e:
        print(f"Ошибка при клике по координатам: {e}")

    return None


def chart_finding(self, arm_window, route_name_from_excel):
    """
    Ищет/создает записи в АРМ Нарядчика на основе переданного названия маршрута.

    :param self: Экземпляр класса, содержащего log и т.д.
    :param arm_window: Окно АРМ Нарядчика.
    :param route_name_from_excel: Название маршрута (значение из столбца A Excel).
    """
    # --- Логика поиска и создания записи ---
    # Используем route_name_from_excel вместо "Пробный МЦД-1"
    input_field = get_input_field(self, arm_window, expected_route_name=route_name_from_excel)

    # Очистка поля и ввод нового значения
    if input_field:
        self.log("Поле ввода названия маршрута найдено")
        input_field.set_text("")  # Очищаем поле
        time.sleep(0.5)  # Небольшая пауза
        # ВАЖНО: Убедиться, что route_name_from_excel безопасен для send_keys
        input_field.type_keys(route_name_from_excel + "{ENTER}", with_spaces=True)
        self.log(f"✅ Введено название маршрута: '{route_name_from_excel}'")
    else:
        self.log("Поле ввода названия маршрута не найдено!")

    # Передаём route_name_from_excel в find_exact_row
    row = find_exact_row(self, arm_window, expected_route_name=route_name_from_excel)
    if row:
        self.log("✅ row is finded!")
        # TODO: Добавьте логику работы со строкой, если она найдена
        # Например, выделение, редактирование и т.д.
        return row  # Или другой индикатор успеха
    else:
        self.log("ℹ️ Row is not found, creating...")
        add_button = get_toolbar_button_by_index(arm_window, 3)
        if not add_button:
            self.log("❌ Кнопка 'Добавить' не найдена!")
            return False

        try:
            # Используем click_input вместо click для обхода COMError
            add_button.click_input()
            self.log("✅ Клик по кнопке 'Добавить' выполнен (click_input).")
        except Exception as click_err:
            self.log(f"❌ Ошибка при клике click_input: {click_err}")
            # Если click_input не сработал, пробуем обычный click как последний шанс
            try:
                add_button.click()
                self.log("✅ Клик по кнопке 'Добавить' выполнен (click).")
            except Exception as click_err2:
                self.log(f"❌ Ошибка при клике click: {click_err2}")
                self.log("❌ Не удалось кликнуть по кнопке 'Добавить' ни одним способом.")
                return False

        try:
            # --- Работа с окном "Новая запись" ---
            # Подключение к приложению
            process_id = arm_window.process_id()
            app_for_new_window = Application(backend="uia").connect(process=process_id)
            main_window = app_for_new_window.window(title="Новая запись", control_type="Window")
            self.log("✅ main_window is went")

            # Ждем, пока окно будет готово
            time.sleep(2)

            # 1. Находим поле для ввода по AutomationId (самый надежный способ)
            # ВАЖНО: Здесь используется arm_window, как в оригинале.
            input_field_new_record = arm_window.child_window(
                auto_id="edtLastName",
                control_type="Edit",
                class_name="WindowsForms10.EDIT.app.0.378734a"
            )
            if input_field_new_record.exists():
                self.log("✅ input_field is inited")

            if not input_field_new_record.exists():
                self.log("❌ Поле ввода не найдено")
                raise Exception("Поле ввода не найдено")

            # 2. Заполняем поле с проверками
            input_field_new_record.set_focus()
            input_field_new_record.set_text("")  # Очищаем поле
            time.sleep(0.3)
            # Используем значение из Excel
            input_field_new_record.type_keys(route_name_from_excel, with_spaces=True)
            self.log(f"✅ Заполнено поле в окне 'Новая запись': '{route_name_from_excel}'")

            # Проверяем что текст введен корректно (закомментировано, как в оригинале)
            # def wait_until(timeout, interval, condition):
            #     end_time = time.time() + timeout
            #     while time.time() < end_time:
            #         if condition():
            #             return True
            #         time.sleep(interval)
            #     return False
            # if not wait_until(5, 0.5, lambda: route_name_from_excel in (input_field_new_record.get_value() or "")):
            #      self.log("⚠️ Текст в поле ввода может быть не полностью установлен.")

            # 3. Находим кнопку "Применить" (по AutomationId)
            apply_button = arm_window.child_window(
                auto_id="btnOk",
                title="Применить",
                control_type="Button"
            )
            if apply_button.exists():
                self.log("✅ apply was")

            if not apply_button.exists():
                self.log("❌ Кнопка 'Применить' не найдена")
                raise Exception("Кнопка 'Применить' не найдена")

            # 4. Нажимаем кнопку с проверками
            apply_button.wait('enabled', timeout=5)
            apply_button.click_input()

            self.log("✅ Операция выполнена успешно")
            # Ждем закрытия окна или обновления основного окна
            main_window.wait_not('visible', timeout=10)
            self.log("✅ Окно 'Новая запись' закрыто.")
            newly_created_row = find_exact_row(self, arm_window, expected_route_name=route_name_from_excel)
            if newly_created_row:
                self.log("✅ Вновь созданная строка найдена после закрытия окна 'Новая запись'.")
                return newly_created_row
            else:
                self.log("⚠️ Вновь созданная строка НЕ НАЙДЕНА после закрытия окна 'Новая запись', но запись была создана.")
                # Можно вернуть True или None, в зависимости от вашей логики обработки
                # Если вернуть None, primary_work пропустит period_schedule
                # Если вернуть True, нужно изменить логику в primary_work
                return True

        except Exception as e:
            self.log(f"❌ Ошибка при работе с окном 'Новая запись': {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            # Дополнительная диагностика при ошибке
            # if 'input_field_new_record' in locals() and input_field_new_record.exists():
            #     self.log(f"ℹ️ Текущее значение поля: {input_field_new_record.get_value()}")
            # if 'main_window' in locals() and main_window.exists():
            #     self.log("ℹ️ Элементы окна 'Новая запись':")
            #     main_window.print_control_identifiers()
            return False


def period_schedule(self, arm_window, row, date):
    """
    :param self: Экземпляр класса логгера/исполнителя.
    :param arm_window: Окно АРМ Нарядчика.
    :param row: Найденная строка (результат find_exact_row).
    :param date: Объект datetime, полученный из Excel.
    """
    # --- Извлечение дня, месяца, года из объекта datetime ---
    day_int = date.day
    month_int = date.month
    year_int = date.year
    # Для использования в полях ввода (строки с ведущими нулями)
    day_str = f"{day_int:02d}"
    month_str = f"{month_int:02d}"
    year_str = f"{year_int}"
    # Для поиска месяца (без ведущего нуля, 1-12)
    month_search_int = month_int
    self.log(f"📅 Полученная дата: {day_str}.{month_str}.{year_str}")

    # --- Остальная логика функции ---
    row_children = row.children()

    # Если потомки есть, кликаем по первому
    if row_children:
        first_cell = row_children[0]
        first_cell.click_input()
        self.log("🖱️ Клик по первой ячейке строки выполнен.")
    else:
        # Если потомков нет, кликаем по самой строке
        row.click_input()
        self.log("🖱️ Клик по самой строке выполнен.")

    period_button = get_toolbar_button_by_index(arm_window, 6)  # Предполагаем, что индекс 6 для "Период"
    if period_button:
        period_button.click_input()
        self.log("🖱️ Клик по кнопке 'Период' выполнен.")
    else:
        self.log("❌ Кнопка 'Период' не найдена!")
        return  # Или обработать ошибку

    # --- Работа с окном "Период действия" ---
    try:
        # Ищем окно "Период действия" как дочернее окно arm_window
        period_window = arm_window.child_window(
            auto_id="FrmMVPSTimetableSched",
            control_type="Window"
        )

        self.log("✅ Окно 'Период действия' найдено.")

        # --- 1. Работа с чекбоксом "Разрешить изменения" ---
        try:
            # Ищем чекбокс "Разрешить изменения" внутри окна "Период действия"
            allow_changes_checkbox = period_window.child_window(
                title="Разрешить изменения",
                control_type="CheckBox"
            )

            # Проверяем, существует ли элемент
            if allow_changes_checkbox.exists(timeout=5):
                self.log("✅ Чекбокс 'Разрешить изменения' найден.")

                # Проверяем текущее состояние через ToggleState (предпочтительный способ)
                try:
                    toggle_state = allow_changes_checkbox.get_toggle_state()
                    is_checked = (toggle_state == 1)  # ToggleState: 1 - On
                    self.log(
                        f"ℹ️ Текущее состояние чекбокса: {'Отмечен' if is_checked else 'Не отмечен'} (ToggleState: {toggle_state})")
                except Exception:
                    # Резервный способ: через LegacyIAccessible.State
                    try:
                        legacy_props = allow_changes_checkbox.legacy_properties()
                        state = legacy_props.get('State', 0)
                        is_checked = bool(state & 0x10)  # CHECKED flag
                        self.log(
                            f"ℹ️ Текущее состояние чекбокса: {'Отмечен' if is_checked else 'Не отмечен'} (Legacy State: {hex(state)})")
                    except Exception:
                        self.log("⚠️ Не удалось проверить состояние чекбокса. Предполагаем 'Не отмечен'.")
                        is_checked = False

                # Если чекбокс НЕ отмечен, кликаем, чтобы отметить
                if not is_checked:
                    self.log("🔄 Устанавливаем чекбокс 'Разрешить изменения'...")
                    allow_changes_checkbox.click_input()
                    self.log("✅ Чекбокс 'Разрешить изменения' установлен.")
                else:
                    self.log("✅ Чекбокс 'Разрешить изменения' уже отмечен.")

            else:
                self.log("❌ Чекбокс 'Разрешить изменения' НЕ НАЙДЕН.")
                # Продолжаем выполнение, возможно, он не обязателен или уже отмечен

        except Exception as checkbox_err:
            self.log(f"❌ Ошибка при работе с чекбоксом 'Разрешить изменения': {checkbox_err}")

        # --- 2. Работа с полем ввода "Год" ---
        try:
            self.log("🔍 Поиск поля ввода 'Год'...")

            # Сначала найдем панель инструментов toolStrip1 внутри period_window
            # Из дампа: AutomationId: "tsMain"
            toolstrip = period_window.child_window(
                auto_id="tsMain",
                control_type="ToolBar"
            )
            if toolstrip.exists():
                self.log("✅ Панель инструментов 'toolStrip1' найдена.")
            else:
                self.log(" Панель инструментов 'toolStrip1' не найдена.")

            # Найдем метку "Год:" (это TextBlock)
            year_label = toolstrip.child_window(
                title="Год:",
                control_type="Text"
            )
            if year_label.exists():
                self.log("✅ Метка 'Год:' найдена.")
            else:
                self.log("Метка 'Год:' не найдена.")

            # Найдем поле ввода, которое находится *после* метки "Год:"
            # Из дампа видно, что поле ввода идет после метки.
            # Можно использовать next_sibling(), если он работает стабильно.
            # Более надежный способ - найти все Edit внутри toolstrip и выбрать нужное.

            # Альтернатива 1: Поиск по соседству (если next_sibling работает)
            # year_input_candidate = year_label.next_sibling(control_type="Edit")

            # Альтернатива 2: Поиск всех Edit и фильтрация по позиции
            edit_fields_in_toolbar = toolstrip.descendants(control_type="Edit")
            year_input = None

            if edit_fields_in_toolbar:
                label_rect = year_label.rectangle()
                # Ищем поле ввода, которое находится правее метки "Год:"
                for edit_field in edit_fields_in_toolbar:
                    try:
                        edit_rect = edit_field.rectangle()
                        # Проверяем, находится ли поле правее метки и примерно на том же уровне по Y
                        if (edit_rect.left > label_rect.right and
                                abs(edit_rect.top - label_rect.top) < 10):  # Допустимая разница по Y
                            year_input = edit_field
                            break
                    except Exception as rect_err:
                        self.log(f"⚠️ Ошибка получения координат поля ввода: {rect_err}")
                        continue

            if year_input:
                self.log("✅ Поле ввода 'Год' найдено.")
                # Очищаем и вводим год
                year_input.set_text("")  # Очищаем
                time.sleep(0.2)  # Небольшая пауза
                year_input.type_keys(year_str)  # Вводим год
                self.log(f"✅ В поле 'Год' введено значение: {year_str}")
            else:
                self.log("❌ Поле ввода 'Год' НЕ НАЙДЕНО.")
                # TODO: Обработка ошибки, если поле ввода года критично

        except Exception as year_input_err:
            self.log(f"❌ Ошибка при работе с полем ввода 'Год': {year_input_err}")
            import traceback
            self.log(traceback.format_exc())

        # --- 3. Работа с календарной таблицей grCalendar ---
        try:
            self.log("🔍 Поиск таблицы календаря 'grCalendar'...")

            # Найдем панель с таблицей (AutomationId: "pnMain")
            calendar_panel = period_window.child_window(
                auto_id="pnMain",
                control_type="Pane"
            )
            if calendar_panel:
                self.log("✅ Панель календаря 'pnMain' найдена.")
            else:
                self.log("Панель календаря 'pnMain' не найдена.")

            # Найдем саму таблицу внутри панели (AutomationId: "grCalendar")
            calendar_table = calendar_panel.child_window(
                auto_id="grCalendar",
                control_type="Table"
            ).wait('exists', timeout=7)
            self.log("✅ Таблица календаря 'grCalendar' найдена.")

            # --- 4. Поиск строки месяца ---
            self.log(f"🔍 Поиск строки для месяца: {month_search_int}...")

            # Словарь для сопоставления номера месяца с названием в Legacy.Value строки
            # Legacy.Value строки выглядит как: "Январь;-;...;-"
            months_map = {
                1: "Январь",
                2: "Февраль",
                3: "Март",
                4: "Апрель",
                5: "Май",
                6: "Июнь",
                7: "Июль",
                8: "Август",
                9: "Сентябрь",
                10: "Октябрь",
                11: "Ноябрь",
                12: "Декабрь"
            }
            target_month_name = months_map.get(month_search_int, "")
            target_month_row = None

            if not target_month_name:
                self.log(f"❌ Неизвестный номер месяца: {month_search_int}")
                raise ValueError(f"Неизвестный номер месяца: {month_search_int}")

            # Получаем все строки таблицы (кроме "Верхняя строка")
            # Ищем дочерние элементы с типом, соответствующим строкам.
            # Из дампа: Name: "Строка 0", ControlType: не указан, Legacy.Role: строка (0x1C)
            # Попробуем искать по имени, начинающемуся с "Строка"
            potential_month_rows = calendar_table.children()  # Получаем всех прямых потомков таблицы

            print(f"   Найдено потенциальных строк в таблице: {len(potential_month_rows)}")

            for i, row_elem in enumerate(potential_month_rows):
                try:
                    row_name = row_elem.window_text()
                    # Проверяем, является ли элемент строкой (но не "Верхняя строка")
                    if row_name.startswith("Строка") and row_name != "Верхняя строка":
                        # Получаем Legacy.Value
                        legacy_props = row_elem.legacy_properties()
                        legacy_value = legacy_props.get('Value', '')
                        print(f"   Проверка строки '{row_name}': Legacy.Value = '{legacy_value}'")

                        # Проверяем, содержит ли Legacy.Value название целевого месяца
                        if target_month_name in legacy_value:
                            target_month_row = row_elem
                            self.log(f"✅ Найдена строка месяца '{target_month_name}': '{row_name}'")
                            break
                except Exception as row_check_err:
                    self.log(f"   ⚠️ Ошибка проверки строки {i}: {row_check_err}")
                    continue

            if not target_month_row:
                self.log(f"❌ Строка для месяца '{target_month_name}' не найдена в таблице.")
                raise Exception(f"Строка месяца '{target_month_name}' не найдена")

            # --- 5. Поиск ячейки дня ---
            self.log(f"🔍 Поиск ячейки для дня: {day_int}...")
            target_day_cell = None

            # Получаем дочерние элементы (ячейки) найденной строки месяца
            day_cells = target_month_row.children(control_type="DataItem")
            # Из дампа: Name: "1 Строка 0", ControlType: DataItem
            self.log(f"   Найдено ячеек в строке месяца: {len(day_cells)}")

            for cell in day_cells:
                try:
                    cell_name = cell.window_text()
                    # Имя ячейки имеет формат "{день} Строка {номер_строки}"
                    # Например: "1 Строка 0", "15 Строка 3"
                    if cell_name.startswith(f"{day_int} Строка"):
                        target_day_cell = cell
                        self.log(f"✅ Найдена ячейка дня {day_int}: '{cell_name}'")
                        break
                except Exception as cell_check_err:
                    self.log(f"   ⚠️ Ошибка проверки ячейки: {cell_check_err}")
                    continue

            if not target_day_cell:
                self.log(f"❌ Ячейка для дня {day_int} не найдена в строке месяца.")
                raise Exception(f"Ячейка дня {day_int} не найдена")

            # --- 6. Клик правой кнопкой мыши по ячейке ---
            self.log(f"🖱️ Клик правой кнопкой мыши по ячейке '{target_day_cell.window_text()}'...")
            # Используем mouse.click из pywinauto для клика правой кнопкой
            # Получаем центральные координаты ячейки
            cell_rect = target_day_cell.rectangle()
            center_x = cell_rect.left + (cell_rect.width() // 2)
            center_y = cell_rect.top + (cell_rect.height() // 2)

            # Клик правой кнопкой мыши
            mouse.click(button='right', coords=(center_x, center_y))
            self.log("✅ Клик правой кнопкой мыши выполнен.")

            # --- 7. Работа с контекстным меню: выбор "Закрепить" ---
            try:
                self.log("🔍 Поиск пункта 'Закрепить' в контекстном меню...")

                # После клика правой кнопкой меню появляется как дочерний элемент period_window.
                # Из дампа: "DropDown" меню -> "Период действия" окно
                # Ищем меню "DropDown" внутри окна "Период действия" (period_window)
                context_menu = period_window.child_window(
                    title="DropDown",  # Имя из дампа
                    control_type="Menu"  # UIA_MenuControlTypeId
                )

                self.log("✅ Контекстное меню 'DropDown' найдено.")

                # Теперь ищем пункт "Закрепить" внутри этого меню
                pin_menu_item = context_menu.child_window(
                    title="Закрепить",
                    control_type="MenuItem"  # UIA_MenuItemControlTypeId
                )

                if pin_menu_item.exists():
                    self.log("✅ Пункт меню 'Закрепить' найден.")

                    # Кликаем левой кнопкой мыши по пункту "Закрепить"
                    # Можно использовать click_input() или invoke() (так как IsInvokePatternAvailable: true)
                    # click_input() обычно надежнее
                    pin_menu_item.click_input()
                    self.log("✅ Клик по пункту меню 'Закрепить' выполнен.")

                else:
                    self.log("❌ Пункт меню 'Закрепить' НЕ НАЙДЕН в контекстном меню.")
                    # TODO: Обработка ошибки, если пункт меню критичен

            except Exception as menu_err:
                self.log(f"❌ Ошибка при работе с контекстным меню: {menu_err}")
                import traceback
                self.log(traceback.format_exc())
                # Если не удалось взаимодействовать с меню, это критично для этой части логики
                return  # Или другая логика обработки ошибки

            # --- 8. Закрытие окна "Период действия" ---
            try:
                self.log("🚪 Попытка закрытия окна 'Период действия'...")
                # Используем period_window_spec, который у нас есть в области видимости
                close_button_spec = period_window_spec.child_window(
                    automation_id="Close", # Используем AutomationId из дампа
                    control_type="Button",
                    title="Закрыть" # Добавим имя для дополнительной уверенности
                )
                # Ждем появления кнопки и получаем Wrapper
                close_button = close_button_spec.wait('exists', timeout=5)
                
                if close_button.exists():
                    close_button.click_input()
                    self.log("✅ Кнопка 'Закрыть' окна 'Период действия' нажата.")
                    # Ждем, пока окно закроется
                    period_window_spec.wait_not('visible', timeout=10)
                    self.log("✅ Окно 'Период действия' закрыто.")
                else:
                    self.log("⚠️ Кнопка 'Закрыть' окна 'Период действия' НЕ НАЙДЕНА.")
                    # Альтернатива: отправить Alt+F4 в окно?
                    # period_window_spec.type_keys("%{F4}") 

            except Exception as close_period_err:
                self.log(f"⚠️ Ошибка при закрытии окна 'Период действия': {close_period_err}")
                # Не критично, продолжаем закрытие основных окон

        except Exception as table_err:
            self.log(f"❌ Ошибка при работе с календарной таблицей: {table_err}")
            import traceback
            self.log(traceback.format_exc())
            return  # Или другая логика обработки ошибки

        self.log("✅ period_schedule (часть с календарем) завершена.")

    except Exception as window_err:
        self.log(f"❌ Критическая ошибка при работе с окном 'Период действия': {window_err}")
        import traceback
        self.log(traceback.format_exc())
        return  # Или другая логика обработки ошибки



def primary_work(self):
    """Основная функция автоматизации."""
    self.log("🚀 Запуск автоматизации...")
    try:
        # --- 1. Парсинг Excel ---
        self.log("📂 Парсинг Excel-файла...")
        try:
            schedule_entries = parse_excel_schedule() # Использует файл по умолчанию
            # Или schedule_entries = parse_excel_schedule("путь/к/вашему/файлу.xlsx")
        except FileNotFoundError as fnf_err:
            self.log(f"❌ {fnf_err}")
            return # Завершаем, если файл не найден
        except Exception as parse_err:
            self.log(f"❌ Ошибка парсинга Excel: {parse_err}")
            import traceback
            self.log(traceback.format_exc())
            return

        if not schedule_entries:
            self.log("⚠️ Из Excel не загружено ни одной записи (возможно, столбцы A или O пусты).")
            return

        self.log(f"✅ Загружено {len(schedule_entries)} записей из Excel.")

        # --- 2. Запуск Client Manager и подключение к АРМ ---
        self.log("🔄 Запуск Client Manager и подключение к АРМ Нарядчика...")
        start_client(self) # Предполагаем, что эта функция не меняется
        arm_window = arm_window_connect(self) # Предполагаем, что эта функция не меняется

        if not arm_window:
            self.log("❌ Не удалось подключиться к АРМ Нарядчика.")
            return

        self.log("✅ Подключение к АРМ Нарядчика выполнено.")

        # --- 3. Обработка каждой записи из Excel ---
        # Проверим, определены ли необходимые функции
        if 'chart_finding' not in globals():
            self.log("❌ Функция chart_finding не найдена.")
            return
        if 'period_schedule' not in globals():
            self.log("❌ Функция period_schedule не найдена.")
            return

        for i, entry in enumerate(schedule_entries):
            # Используем значение из СТОЛБЦА A (первый столбец) текущей записи
            route_name = str(entry.col_A) if entry.col_A is not None else ""

            # --- Используем значение из СТОЛБЦА O (15-й столбец) - дата ---
            raw_date_value = entry.col_O
            date_object = None
            if raw_date_value is not None and isinstance(raw_date_value, datetime.datetime): # Исправлено: datetime.datetime
                date_object = raw_date_value
                date_log_str = date_object.strftime('%d.%m.%Y')
            elif raw_date_value is not None:
                try:
                    # Попробуем распарсить, если это строка
                    date_object = datetime.datetime.strptime(str(raw_date_value), '%d.%m.%Y') # Исправлено: datetime.datetime.strptime
                    date_log_str = str(raw_date_value)
                except ValueError:
                    self.log(f"⚠️ Запись {i + 1}: Невозможно распарсить дату '{raw_date_value}'. Пропущена.")
                    continue
            else:
                self.log(f"⚠️ Запись {i + 1}: Пустая дата в столбце O. Пропущена.")
                continue
            # --- Конец извлечения/проверки даты ---

            if not route_name:
                self.log(f"⚠️ Запись {i + 1}/{len(schedule_entries)} пропущена: пустое значение в столбце A.")
                continue

            self.log(f"--- Обработка записи {i + 1}/{len(schedule_entries)}: '{route_name}' (Дата: {date_log_str}) ---")

            # --- Шаг 1: Вызов chart_finding ---
            try:
                # Предполагаем сигнатуру: chart_finding(self, arm_window, route_name_from_excel)
                row = chart_finding(self, arm_window, route_name_from_excel=route_name)
                if row:
                    self.log(f"✅ chart_finding для '{route_name}' выполнена успешно. Получен объект строки.")
                else:
                    self.log(f"⚠️ chart_finding для '{route_name}' завершилась без результата (строка не найдена/создана?). Пропущена запись.")
                    continue # Пропускаем period_schedule для этой итерации
            except Exception as e:
                self.log(f"❌ Критическая ошибка в chart_finding для '{route_name}': {e}")
                import traceback
                self.log(traceback.format_exc())
                continue # Пропускаем period_schedule для этой итерации при ошибке chart_finding

            # --- Шаг 2: Вызов period_schedule (только если row и date_object были получены) ---
            if row and date_object: # Проверяем, что и row не None, и date_object не None
                try:
                    # Вызываем функцию period_schedule
                    # Предполагаем сигнатуру: def period_schedule(self, arm_window, row, date):
                    period_schedule(self, arm_window, row, date_object)
                    self.log(f"✅ period_schedule для '{route_name}' с датой '{date_log_str}' выполнена.")
                except Exception as e:
                    self.log(f"❌ Ошибка в period_schedule для '{route_name}' с датой '{date_log_str}': {e}")
                    import traceback
                    self.log(traceback.format_exc())
                    # Можно решить, продолжать ли выполнение или прервать
            elif not date_object:
                self.log(f"ℹ️ period_schedule для '{route_name}' не вызывается: объект даты не был создан.")
            else: # row is None
                self.log(f"ℹ️ period_schedule для '{route_name}' не вызывается, так как row не был получен.")

            # Опционально: добавить небольшую паузу между итерациями
            # import time
            # time.sleep(1) # Например, 1 секунда

        self.log("🎉 Все записи из Excel обработаны (или попытались обработать).")

        # --- 4. Закрытие приложений ---
        self.log("🚪 Начало процедуры закрытия приложений...")
        
        # --- 4.1 Закрытие "АРМ Нарядчика" ---
        try:
            # Проверим, что arm_window все еще существует и доступен
            if 'arm_window' in locals() and arm_window and arm_window.exists():
                self.log("🚪 Попытка закрытия 'АРМ Нарядчика'...")
                
                # Получаем спецификацию главного окна АРМ Нарядчика
                # Нам нужно получить родительское окно для кнопки закрытия
                # Лучше искать само окно "АРМ Нарядчика" и затем кнопку в нем
                arm_window_title = arm_window.window_text()
                main_arm_window_spec = Application(backend="uia").connect(process=arm_window.process_id()).window(title=arm_window_title)
                
                # Ищем кнопку "Закрыть" в главном окне АРМ Нарядчика
                arm_close_button_spec = main_arm_window_spec.child_window(
                    automation_id="Close", # Используем AutomationId из дампа
                    control_type="Button",
                    title="Закрыть"
                )
                
                if arm_close_button_spec.exists(timeout=5):
                    arm_close_button = arm_close_button_spec.wrapper_object()
                    arm_close_button.click_input()
                    self.log("✅ Кнопка 'Закрыть' окна 'АРМ Нарядчика' нажата.")
                    
                    # Ждем, пока окно закроется
                    try:
                        main_arm_window_spec.wait_not('visible', timeout=15)
                        self.log("✅ Окно 'АРМ Нарядчика' закрыто.")
                    except:
                        self.log("⚠️ Окно 'АРМ Нарядчика' не закрылось за отведенное время (возможно, появился диалог подтверждения).")
                        # Если окно не закрылось, попробуем Alt+F4 как запасной вариант
                        try:
                            main_arm_window_spec.type_keys("%{F4}")
                            main_arm_window_spec.wait_not('visible', timeout=5)
                            self.log("✅ Окно 'АРМ Нарядчика' закрыто после Alt+F4.")
                        except:
                            self.log("⚠️ Закрытие 'АРМ Нарядчика' с помощью Alt+F4 также не удалось.")
                else:
                    self.log("⚠️ Кнопка 'Закрыть' окна 'АРМ Нарядчика' НЕ НАЙДЕНА. Пробуем Alt+F4.")
                    # Альтернатива: отправить Alt+F4 в главное окно
                    try:
                        main_arm_window_spec.type_keys("%{F4}")
                        main_arm_window_spec.wait_not('visible', timeout=10)
                        self.log("✅ Окно 'АРМ Нарядчика' закрыто после Alt+F4.")
                    except:
                         self.log("⚠️ Закрытие 'АРМ Нарядчика' с помощью Alt+F4 также не удалось.")
                         
            else:
                self.log("ℹ️ Окно 'АРМ Нарядчика' уже закрыто или недоступно.")
        except Exception as close_arm_err:
            self.log(f"⚠️ Ошибка при закрытии 'АРМ Нарядчика': {close_arm_err}")

        # --- 4.2 Закрытие "Управление клиентом" ---
        try:
            self.log("🚪 Попытка закрытия 'Управление клиентом'...")
            # Нужно подключиться к уже запущенному приложению "Управление клиентом"
            client_manager_app = Application(backend="uia").connect(title_re=".*Управление клиентом.*", timeout=5)
            client_manager_window = client_manager_app.window(title_re=".*Управление клиентом.*")
            
            if client_manager_window.exists():
                # Ищем кнопку "Закрыть" в окне "Управление клиентом"
                client_close_button_spec = client_manager_window.child_window(
                    automation_id="Close", # Используем AutomationId из дампа
                    control_type="Button",
                    title="Закрыть"
                )
                
                if client_close_button_spec.exists(timeout=5):
                    client_close_button = client_close_button_spec.wrapper_object()
                    client_close_button.click_input()
                    self.log("✅ Кнопка 'Закрыть' окна 'Управление клиентом' нажата.")
                    
                    # Ждем, пока окно закроется
                    try:
                        client_manager_window.wait_not('visible', timeout=10)
                        self.log("✅ Окно 'Управление клиентом' закрыто.")
                    except:
                        self.log("⚠️ Окно 'Управление клиентом' не закрылось за отведенное время.")
                        # Пробуем Alt+F4
                        try:
                            client_manager_window.type_keys("%{F4}")
                            client_manager_window.wait_not('visible', timeout=5)
                            self.log("✅ Окно 'Управление клиентом' закрыто после Alt+F4.")
                        except:
                            self.log("⚠️ Закрытие 'Управление клиентом' с помощью Alt+F4 также не удалось.")
                else:
                    self.log("⚠️ Кнопка 'Закрыть' окна 'Управление клиентом' НЕ НАЙДЕНА. Пробуем Alt+F4.")
                    # Альтернатива: отправить Alt+F4
                    try:
                        client_manager_window.type_keys("%{F4}")
                        client_manager_window.wait_not('visible', timeout=10)
                        self.log("✅ Окно 'Управление клиентом' закрыто после Alt+F4.")
                    except:
                        self.log("⚠️ Закрытие 'Управление клиентом' с помощью Alt+F4 также не удалось.")
                        
            else:
                 self.log("ℹ️ Окно 'Управление клиентом' не найдено при попытке закрытия.")
                 
        except Exception as connect_close_client_err:
            self.log(f"ℹ️ Не удалось подключиться или закрыть 'Управление клиентом': {connect_close_client_err}. Возможно, оно уже закрыто.")
            
        self.log("🎉 Основные приложения закрыты (или попытались закрыть).")

    except Exception as e:
        self.log(f"❌ Критическая ошибка в primary_work: {e}")
        import traceback
        self.log(traceback.format_exc())


def open_in_paint(app_instance):
    """Открытие свойств изображения в Paint"""
    try:
        # Проверка существования файла
        if not os.path.exists(app_instance.IMAGE_PATH):
            app_instance.log(f"❌ Файл не найден: {app_instance.IMAGE_PATH}")
            return

        # Закрытие предыдущих экземпляров Paint
        for proc in psutil.process_iter():
            if proc.name() == app_instance.PAINT_EXE:
                proc.kill()

        app_instance.log("🎨 Запускаем Paint с изображением...")

        # Настройка таймаутов для более стабильной работы
        timings.Timings.fast()

        # Открытие изображения в Paint
        paint_app = Application(backend="uia").start(f'{app_instance.PAINT_EXE} "{app_instance.IMAGE_PATH}"')
        time.sleep(2)  # Ожидание запуска

        try:
            # Подключение к основному окну Paint
            paint_window = paint_app.window(title_re=".* - Paint")
            app_instance.log("✅ Paint успешно запущен")

            # 1. Открываем меню "Файл"
            app_instance.log("📂 Открываем меню 'Файл'...")
            paint_window.menu_select("Файл")
            time.sleep(0.5)

            # 2. Выбираем "Свойства изображения"
            app_instance.log("⚙️ Открываем свойства изображения...")
            paint_window.menu_select("Файл -> Свойства изображения")
            time.sleep(1)

            # 3. Получаем информацию из диалога свойств
            props_dialog = paint_app.window(title="Свойства изображения")

            # Получаем размеры
            width = props_dialog.child_window(auto_id="1148", control_type="Edit").window_text()
            height = props_dialog.child_window(auto_id="1149", control_type="Edit").window_text()
            units = props_dialog.child_window(auto_id="1152", control_type="ComboBox").selected_text()

            app_instance.log(f"📏 Размер изображения: {width}x{height} {units}")

            # 4. Закрываем диалог свойств
            props_dialog.cancel.click()
            time.sleep(0.5)

            # 5. Закрываем Paint
            paint_window.close()
            app_instance.log("🖌️ Работа с Paint завершена")

        except ElementNotFoundError as e:
            app_instance.log(f"⚠️ Ошибка поиска элемента: {str(e)}")
            # Попробуем альтернативный способ через горячие клавиши
            try:
                send_keys('%Ф')  # Alt+Ф (меню Файл)
                send_keys('С')  # Свойства изображения
                time.sleep(1)
                app_instance.log("ℹ️ Использован альтернативный метод открытия свойств")
            except Exception as kbd_ex:
                app_instance.log(f"⚠️ Ошибка клавиатурного ввода: {str(kbd_ex)}")

    except Exception as e:
        app_instance.log(f"❌ Критическая ошибка: {str(e)}")
