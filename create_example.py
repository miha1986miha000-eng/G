#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для создания примера Excel файла
"""

import pandas as pd
from openpyxl import Workbook
from datetime import datetime, timedelta

def create_example_excel():
    """Создает пример Excel файла с данными"""

    # Создаем данные для примера
    data = {
        'ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        'Название': [
            'Товар А',
            'Товар Б',
            'Товар В',
            'Товар Г',
            'Товар Д',
            'Товар Е',
            'Товар Ж',
            'Товар З',
            'Товар И',
            'Товар К',
            'Товар А',  # Дубликат
            'Товар Л'
        ],
        'Количество': [100, 250, 75, 320, 150, 90, 200, 180, 110, 95, 100, 145],
        'Цена': [1500.50, 2300.00, 890.75, 4500.25, 1200.00, 3300.50, 1800.00, 2100.00, 950.00, 1650.00, 1500.50, 2800.00],
        'Статус': ['Активен', 'Активен', 'Неактивен', 'Активен', 'Активен', 'Неактивен', 'Активен', 'Активен', 'Неактивен', 'Активен', 'Активен', 'Активен'],
        'Дата': [
            (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
            for i in range(12)
        ]
    }

    # Создаем DataFrame
    df = pd.DataFrame(data)

    # Добавляем несколько пустых строк
    empty_rows = pd.DataFrame([[None] * len(df.columns)] * 2, columns=df.columns)
    df = pd.concat([df.iloc[:5], empty_rows, df.iloc[5:]], ignore_index=True)

    # Сохраняем в Excel
    output_file = 'example.xlsx'
    df.to_excel(output_file, index=False)

    print(f"Пример Excel файла создан: {output_file}")
    print(f"Файл содержит:")
    print(f"  - {len(df)} строк (включая пустые)")
    print(f"  - {len(df.columns)} колонок")
    print(f"  - Дубликаты данных")
    print(f"  - Пустые строки")
    print(f"\nИспользуйте этот файл для тестирования Excel Data Formatter!")

if __name__ == "__main__":
    create_example_excel()
