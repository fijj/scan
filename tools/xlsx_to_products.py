"""Конвертирует выгрузку остатков (xlsx) в products.js для приложения инвентаризации.

Использование:
    python tools/xlsx_to_products.py "C:/path/PRC_xxx.xlsx" [products.js]

Ожидаемые колонки (по заголовку): Артикул, Наименование, Доступно, Штрихкод.
Строки без штрихкода пропускаются. Пустое «Доступно» считается нулём.
"""
import json
import re
import sys

import openpyxl

REQUIRED = {'article': 'Артикул', 'name': 'Наименование', 'stock': 'Доступно', 'barcode': 'Штрихкод'}


def parse_qty(value):
    """'1 234,000' / '48,000' / 12 / None -> число (int, если целое)."""
    if value is None or value == '':
        return 0
    if isinstance(value, (int, float)):
        num = float(value)
    else:
        text = re.sub(r'[\s\u00a0]', '', str(value)).replace(',', '.')
        num = float(text) if text else 0.0
    return int(num) if num.is_integer() else num


def main():
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else 'products.js'

    ws = openpyxl.load_workbook(src, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else '' for h in next(rows)]
    try:
        idx = {key: header.index(title) for key, title in REQUIRED.items()}
    except ValueError as e:
        sys.exit(f'Не найдена колонка: {e}. Заголовок файла: {header}')

    items, seen = [], set()
    no_barcode, duplicates = [], []
    for row in rows:
        article = str(row[idx['article']] or '').strip()
        barcode = str(row[idx['barcode']] or '').strip()
        if barcode.endswith('.0'):
            barcode = barcode[:-2]
        if not barcode:
            if article:
                no_barcode.append(article)
            continue
        if barcode in seen:
            duplicates.append(barcode)
            continue
        seen.add(barcode)
        items.append({
            'article': article,
            'name': str(row[idx['name']] or '').strip(),
            'barcode': barcode,
            'stock': parse_qty(row[idx['stock']]),
        })

    with open(dst, 'w', encoding='utf-8') as f:
        f.write(f'// Сгенерировано из {src.replace(chr(92), "/").split("/")[-1]}\n')
        f.write('const PRODUCTS = [\n')
        f.write(',\n'.join('  ' + json.dumps(i, ensure_ascii=False) for i in items))
        f.write('\n];\n')

    with_stock = sum(1 for i in items if i['stock'])
    print(f'Товаров: {len(items)}, с остатком: {with_stock}, штук всего: {sum(i["stock"] for i in items)}')
    print(f'Без штрихкода пропущено: {len(no_barcode)} {no_barcode}')
    print(f'Дубликаты штрихкодов: {len(duplicates)} {duplicates}')


if __name__ == '__main__':
    main()
