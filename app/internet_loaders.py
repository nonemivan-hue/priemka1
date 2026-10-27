# -*- coding: utf-8 -*-
"""Загрузка справочников из сети Интернет.

1. Страны происхождения (ОКМ / Alpha-2 → цифровой код) — берём открытый
   список ISO 3166 с GitHub (countries.json).
2. Единицы измерения (ОКЕИ) — открытый список ОКЕИ (OKI csv/json).
3. ФИАС (адресный классификатор) — официальная загрузка с fias.nalog.ru
   очень велика, поэтому программа умеет подгружать компактный набор
   адресов из открытого датасета и позволяет импорт ФИАС из файла XML/CSV.

Все функции возвращают списки словарей; при недоступности сети используется
локальный кэш (data/cache_*.json), а при его отсутствии — встроенный
минимальный набор данных.
"""
import json
import os
import socket
import urllib.error
import urllib.request

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(ROOT_DIR, 'data')

socket.setdefaulttimeout(15)

# --- Встроенный минимум (используется, если нет ни сети, ни кэша) ---------
MINIMAL_COUNTRIES = [
    {"kod": "643", "naimovanie": "Россия"},
    {"kod": "112", "naimovanie": "Беларусь"},
    {"kod": "398", "naimovanie": "Казахстан"},
    {"kod": "156", "naimovanie": "Китай"},
    {"kod": "792", "naimovanie": "Турция"},
    {"kod": "276", "naimovanie": "Германия"},
    {"kod": "840", "naimovanie": "США"},
    {"kod": "410", "naimovanie": "Республика Корея"},
    {"kod": "392", "naimovanie": "Япония"},
    {"kod": "250", "naimovanie": "Франция"},
]

MINIMAL_OKOI = [
    {"kod": "354", "obozachenie": "шт.", "naimovanie": "Штука"},
    {"kod": "362", "obozachenie": "мес.", "naimovanie": "Месяц"},
    {"kod": "796", "obozachenie": "усл. ед.", "naimovanie": "Условная единица"},
    {"kod": "876", "obozachenie": "усл. ед.", "naimovanie": "Услуга"},
    {"kod": "113", "obozachenie": "кг", "naimovanie": "Килограмм"},
    {"kod": "066", "obozachenie": "м", "naimovanie": "Метр"},
    {"kod": "110", "obozachenie": "кв. м", "naimovanie": "Квадратный метр"},
    {"kod": "377", "obozachenie": "компл.", "naimovanie": "Комплект"},
    {"kod": "612", "obozachenie": "тыс. шт.", "naimovanie": "Тысяча штук"},
    {"kod": "577", "obozachenie": "л", "naimovanie": "Литр"},
]

# ISO alpha-2 -> цифровой код (для конвертации стран из интернета)
ISO2_TO_NUM = {
    "AD": "020", "AE": "784", "AF": "971", "AL": "008", "AM": "051", "AO": "024",
    "AR": "032", "AT": "040", "AU": "036", "AZ": "031", "BA": "070", "BD": "050",
    "BE": "056", "BF": "854", "BG": "100", "BH": "048", "BI": "108", "BJ": "204",
    "BN": "096", "BO": "068", "BR": "076", "BS": "044", "BT": "064", "BW": "072",
    "BY": "112", "BZ": "084", "CA": "124", "CD": "180", "CF": "140", "CG": "178",
    "CH": "756", "CI": "384", "CL": "152", "CM": "120", "CN": "156", "CO": "170",
    "CR": "188", "CU": "192", "CY": "196", "CZ": "203", "DE": "276", "DJ": "262",
    "DK": "208", "DO": "214", "DZ": "012", "EC": "218", "EE": "233", "EG": "818",
    "ER": "232", "ES": "724", "ET": "231", "FI": "246", "FJ": "242", "FR": "250",
    "GB": "826", "GE": "268", "GH": "288", "GM": "270", "GN": "324", "GR": "300",
    "GT": "320", "GY": "328", "HN": "340", "HR": "191", "HT": "332", "HU": "348",
    "ID": "360", "IE": "372", "IL": "376", "IN": "356", "IQ": "368", "IR": "364",
    "IS": "352", "IT": "380", "JM": "388", "JO": "400", "JP": "392", "KE": "404",
    "KG": "417", "KP": "408", "KR": "410", "KW": "414", "KZ": "398", "LA": "418",
    "LB": "422", "LI": "438", "LK": "144", "LR": "430", "LS": "426", "LT": "440",
    "LU": "442", "LV": "428", "LY": "434", "MA": "504", "MD": "498", "ME": "499",
    "MG": "450", "MK": "807", "ML": "466", "MM": "104", "MN": "496", "MR": "478",
    "MT": "470", "MU": "480", "MV": "462", "MW": "454", "MX": "484", "MY": "458",
    "MZ": "508", "NA": "516", "NE": "562", "NG": "566", "NI": "558", "NL": "528",
    "NO": "578", "NP": "524", "NR": "520", "NZ": "554", "OM": "512", "PA": "591",
    "PE": "604", "PG": "598", "PH": "608", "PK": "586", "PL": "616", "PT": "620",
    "PY": "600", "QA": "634", "RO": "642", "RS": "688", "RU": "643", "SA": "682",
    "SD": "729", "SE": "752", "SG": "702", "SI": "705", "SK": "703", "SL": "694",
    "SM": "674", "SN": "686", "SO": "706", "SR": "740", "SS": "728", "SV": "222",
    "SY": "760", "SZ": "748", "TD": "148", "TG": "768", "TH": "764", "TJ": "762",
    "TM": "795", "TN": "788", "TR": "792", "TT": "780", "TW": "158", "TZ": "834",
    "UA": "804", "UG": "800", "US": "840", "UY": "858", "UZ": "860", "VE": "862",
    "VN": "704", "VU": "548", "YE": "887", "ZA": "710", "ZM": "894", "ZW": "716",
}


def _fetch_json(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode('utf-8'))


def _cache_path(name):
    return os.path.join(CACHE_DIR, 'cache_%s.json' % name)


def _read_cache(name):
    p = _cache_path(name)
    if os.path.exists(p):
        try:
            with open(p, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None
    return None


def _write_cache(name, data):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(_cache_path(name), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)


def load_countries():
    """Страны происхождения: цифровой код + наименование (из интернета)."""
    urls = [
        "https://raw.githubusercontent.com/mledoze/countries/master/dist/countries.json",
    ]
    for u in urls:
        try:
            data = _fetch_json(u)
            result = []
            for c in data:
                name_ru = (c.get('translations', {}).get('rus', {}) or {}).get('common') \
                          or (c.get('translations', {}).get('rus') or '')
                if not name_ru:
                    name_ru = c.get('name', {}).get('common', '')
                num = c.get('numeric')
                if num is None:
                    continue
                result.append({"kod": str(num).zfill(3), "naimovanie": name_ru})
            if result:
                _write_cache('countries', result)
                return result, "загружено из Интернета (ОКМ/ISO-3166): %d стран" % len(result)
        except Exception:
            continue
    cached = _read_cache('countries')
    if cached:
        return cached, "взято из локального кэша предыдущей загрузки: %d стран" % len(cached)
    return list(MINIMAL_COUNTRIES), "сеть недоступна, загружен минимальный набор (%d стран)" % len(MINIMAL_COUNTRIES)


def load_okoi():
    """Единицы измерения ОКЕИ (из интернета, с фолбэками и кэшем)."""
    # Вариант 1: реестр единиц измерения КЭРП (ГНС)
    try:
        url = ("https://egrp.gnc.ru/api/v1/okoi/list?"
               "page=1&pagesize=1000")
        data = _fetch_json(url)
        items = data.get('results') or data.get('data') or []
        result = []
        for it in items:
            kod = str(it.get('code', it.get('kod', ''))).strip()
            oboz = (it.get('oke') or it.get('conditional_designation')
                    or it.get('oboznachenie') or '').strip()
            name = (it.get('name') or it.get('title') or '').strip()
            if kod and oboz:
                result.append({"kod": kod.zfill(3), "obozachenie": oboz, "naimovanie": name})
        if result:
            _write_cache('okoi', result)
            return result, "загружено из Интернета (ОКЕИ ГНС): %d единиц" % len(result)
    except Exception:
        pass
    cached = _read_cache('okoi')
    if cached:
        return cached, "взято из локального кэша предыдущей загрузки: %d единиц" % len(cached)
    return list(MINIMAL_OKOI), "сеть недоступна, загружен минимальный набор ОКЕИ (%d единиц)" % len(MINIMAL_OKOI)


# Источники открытых данных ФИАС/ГАР (проверены, доступны без регистрации).
FIAS_SOURCES = [
    {
        "name": "адреса субъектов РФ (ФИАС)",
        "urls": [
            "https://raw.githubusercontent.com/asakasinsky/russia.json/master/json/russia.subjects.json",
            "https://cdn.jsdelivr.net/gh/asakasinsky/russia.json@master/json/russia.subjects.json",
        ],
        "mode": "subjects",
    },
    {
        "name": "населённые пункты и районы РФ (ФИАС)",
        "urls": [
            "https://raw.githubusercontent.com/asakasinsky/russia.json/master/json/russia.locality.json",
            "https://cdn.jsdelivr.net/gh/asakasinsky/russia.json@master/json/russia.locality.json",
        ],
        "mode": "locality",
        "limit": 3000,
    },
]


def _fetch_bytes(urls, timeout=20):
    """Пробует загрузить URL по очереди; возвращает (bytes, url) или бросает
    последнее исключение сети."""
    last_err = None
    for u in urls:
        try:
            req = urllib.request.Request(u, headers={'User-Agent': 'Mozilla/5.0'})
            return urllib.request.urlopen(req, timeout=timeout).read(), u
        except urllib.error.HTTPError as e:
            last_err = e      # источник недоступен — пробуем следующий
        except (urllib.error.URLError, socket.timeout, OSError) as e:
            last_err = e
    raise last_err or RuntimeError('нет источников')


def _parse_fias_subjects(raw):
    data = json.loads(raw.decode('utf-8', errors='ignore'))
    items = data.get('data') if isinstance(data, dict) else data
    result = []
    for it in items or []:
        name = str(it.get('name', '')).strip()
        if not name:
            continue
        code = str(it.get('regioncode', '')).strip()
        result.append({
            "aoid": str(it.get('aoguid', '')),
            "status": "Действующий",
            "indeks": "",
            "kray_oblast": "%s (код %s)" % (name, code) if code else name,
            "rayon": "", "gorod": "", "naselenny_punkt": "",
            "ulica": "", "dom": "",
        })
    return result


def _parse_fias_locality(raw, limit=3000):
    data = json.loads(raw.decode('utf-8', errors='ignore'))
    items = data.get('data') if isinstance(data, dict) else data
    result = []
    for it in (items or [])[:limit]:
        name = str(it.get('name', '')).strip()
        if not name:
            continue
        parts = name.split(None, 1)
        value = parts[0] if parts else name
        vtype = parts[1] if len(parts) > 1 else ''
        is_city = vtype.lower().startswith(('г', 'город'))
        code = str(it.get('regioncode', '')).strip()
        result.append({
            "aoid": str(it.get('aoguid', '')),
            "status": "Действующий",
            "indeks": "",
            "kray_oblast": "код региона %s" % code if code else "",
            "rayon": "" if is_city else value,
            "gorod": value if is_city else "",
            "naselenny_punkt": "",
            "ulica": "", "dom": "",
        })
    return result


def load_fias_sample(region_code=None):
    """Адресный классификатор (ФИАС).

    Программа подгружает открытые данные ФИАС/ГАР (субъекты РФ, затем
    населённые пункты и районы) с нескольких зеркал; при недоступности сети —
    локальный кэш предыдущей загрузки или встроенный минимум. Полный экспорт
    ФИАС (fias.nalog.ru / ГАР, XML) можно импортировать в справочник файлом
    на странице «Адресный классификатор».
    """
    errors = []
    for src in FIAS_SOURCES:
        try:
            raw, _u = _fetch_bytes(src['urls'])
            if src['mode'] == 'subjects':
                result = _parse_fias_subjects(raw)
            else:
                result = _parse_fias_locality(raw, src.get('limit', 3000))
            if region_code:
                result = [r for r in result if region_code in str(r.get('kray_oblast', ''))] \
                    or result
            if result:
                _write_cache('fias', result)
                return result, "загружено из Интернета (%s): %d записей" % (src['name'], len(result))
        except Exception as e:
            errors.append("%s: %s" % (src['name'], str(e)[:60]))
    cached = _read_cache('fias')
    if cached:
        return cached, ("сеть недоступна (%s); взято из локального кэша предыдущей "
                        "загрузки: %d адресов" % ('; '.join(errors), len(cached)))
    minimal = [{
        "aoid": "", "status": "Действующий", "indeks": "",
        "kray_oblast": "Краснодарский край", "rayon": "",
        "gorod": "Геленджик", "naselenny_punkt": "",
        "ulica": "Курзальная", "dom": "12",
    }]
    return minimal, ("сеть недоступна (%s), загружен минимальный набор адресов; "
                     "полный ФИАС можно импортировать файлом XML/CSV" % '; '.join(errors))
