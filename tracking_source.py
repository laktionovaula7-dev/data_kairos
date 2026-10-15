"""Read the flattened production/transit workbook without changing source values."""
import datetime
import re
import openpyxl


def date_iso(value):
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, str):
        try:
            return datetime.datetime.strptime(value.strip()[:10], '%d.%m.%Y').strftime('%Y-%m-%d')
        except ValueError:
            return None
    return None


def load_normalized_tracking(path, classify):
    result = []
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        header = None
        for line, row in enumerate(wb.active.iter_rows(values_only=True), 1):
            if row[0] == 'Склад назначения' and row[1] == 'Статус по источнику':
                header = line
                continue
            if not header or line <= header or row[1] not in ('В пути', 'В производстве') or not row[2]:
                continue
            if not isinstance(row[4], (int, float)) or row[4] <= 0:
                continue
            destination = ('Владивосток' if 'ВЛ' in row[0] else 'Москва' if 'МСК' in row[0]
                           else 'Новосибирск' if 'Новосибирск' in row[0] else row[0])
            article, brand, category = classify(row[2])
            doc = row[10] or ''
            match = re.search(r'от (\d{2}\.\d{2}\.20\d{2})', doc)
            sea, port, station, raw_eta = map(date_iso, (row[6], row[7], row[8], row[5]))
            result.append(dict(name=row[2], art=article or '', brand=brand, cat=category,
                               qty=row[4], unit=row[3], status=row[1], dest=destination, wh=row[11],
                               order=match.group(1) if match else '', orderDoc=doc,
                               sea=sea, port=port, ktk=station, eta=raw_eta if sea else None,
                               sourceEta=raw_eta, etaUnconfirmed=not bool(sea), sourceRows=[row[12]]))
        if not header:
            raise ValueError('Не найден заголовок таблицы производства и поставок')
    finally:
        wb.close()
    return result
