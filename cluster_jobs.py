#!/usr/bin/env python3
"""
Кластеризация 225 джобов из jobs_master.csv:
- Шаг 1: разнести по 11 топ-блокам из CLAUDE.md
- Шаг 2: внутри блока — фаззи-кластеризация по семантической близости (TF-IDF + cosine)
- Шаг 3: для каждого кластера агрегировать: канон want_short + frequency + segments + sources

Выход:
- jobs_clustered.csv — новая таблица
- clustering_report.md — что попало в какой блок, что не влезло
"""

import csv
import re
from pathlib import Path
from collections import defaultdict, Counter
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

BASE = Path("/Users/romansemenov/Desktop/Cloud Cod/Knowledge/jobs_master_table")
IN_CSV = BASE / "jobs_master.csv"

# 11 блоков из CLAUDE.md + ключевые слова для классификации.
# Слова — корни (для русского стемминга подойдут).
BLOCKS = {
    "Финансы и учёт": [
        "финанс", "прибыль", "выручк", "касс", "оплат", "плате", "счёт", "счета", "доход",
        "расход", "бюджет", "рентабельност", "реализац", "дебиторк", "ндс", "налог",
        "отчётност", "продаж", "товар", "склад", "стоимост", "цен", "цена", "рекуррент",
        "подписк", "абонемент", "сертификат", "скидк", "тариф", "комисс", "пакет", "акци",
        "юкасс", "tco", "доплат", "юр", "договор", "оферт", "капитализац", "дивиденд",
    ],
    "Персонал и зарплаты": [
        "зарплат", "тренер", "админ", "сотрудник", "kpi", "процент", "ставк", "ставка",
        "мотивац", "найм", "обуч", "сменщ", "адаптац", "штат", "аренд", "график работ",
        "роль", "доступ сотрудник", "распределен", "договорённост", "увольнен",
        "ндфл", "перевод", "коуч", "супервайз", "стандарт", "регламент",
    ],
    "Клиенты и удержание": [
        "удержан", "отток", "возврат", "верн", "лояльност", "ltv", "сегментац",
        "клиент", "база", "анкет", "карточк", "контакт", "продлен", "новые клиент",
        "ушедш", "засыпа", "спящ", "повторн", "ретеншн", "программ лояльност",
        "бонус", "промокод", "приведи друг", "связ родител", "ребёнк", "ребенк", "семейн", "детск",
        "отзыв", "обратная связь",
    ],
    "Расписание и запись": [
        "расписан", "запис", "слот", "групповы", "персональн", "бронирован", "клас",
        "тренировк", "занят", "посещен", "повторяющ", "лист ожидан", "пересеч",
        "вместимост", "загруж", "пик", "копировать расписан",
    ],
    "Коммуникации и маркетинг": [
        "рассылк", "сообщен", "коммуникац", "ватсап", "whatsapp", "телеграм", "telegram",
        "макс", "max", "messenger", "инстаграм", "instagram", "вконтакт", "vk",
        "звонк", "ip-телефони", "телефони", "лид", "лидоген", "заявк", "воронк",
        "маркетинг", "реклам", "контекст", "таргет", "smm", "соцсет", "источник трафик",
        "флаер", "промоутер", "штендер", "виндер", "вывеск", "посадк", "кросс",
        "офер", "офферт", "промо", "первое сообщен", "триггерн", "автомат рассылк",
    ],
    "Аналитика и управление": [
        "аналитик", "дашборд", "отчёт", "отчет", "статистик", "ki", "kpi", "бизнес-показ",
        "обзор", "сравнен", "телеграм-бот", "tg-бот", "выгруз", "excel", "финрез",
        "финмодель", "финансовая модель", "сводк", "управленческ", "контроль показател",
        "стратеги", "целеполаган", "планирован",
    ],
    "СКУД и доступ": [
        "скуд", "сигур", "pocketkey", "pocket k", "macwox", "qr", "qr-код", "вход в зал",
        "контроль доступ", "домофон", "турникет", "автономн вход", "kyc",
        "видеоналит", "видеоналитик", "фрод", "верификац", "идентификац",
        "несанкционирован", "проход",
    ],
    "Операционка / не быть узким местом": [
        "узкое место", "ручн", "минимизировать", "автоматизирова", "автоматизац",
        "делегиров", "контрол админ", "ресепшен", "не вмешива", "без моего", "освободи",
        "лично", "вовлеч", "операционк", "снизить нагрузк", "управлять студ",
        "управлять клуб", "управлять сет", "управлять бизнес", "масштабир", "масштаб",
        "стать сетью", "построить", "запустить и развить", "управление компани",
        "разрозненн", "хаос", "единая систем", "одно окно", "одном окне",
    ],
    "Запуск и переход": [
        "запустить", "запуск", "открыт", "новый клуб", "новый зал", "переход",
        "перенос баз", "перенос данных", "миграц", "обучен сотрудник", "внедрен",
        "франшиз", "франчайз", "партнёр", "партнер", "выход из франшиз", "выбор cra",
        "выбор crm", "выбор систем", "стартап", "новый бизнес", "сначала",
    ],
    "Клиентский опыт и приложение": [
        "приложен", "мпк", "мобильн", "мобайл", "брендир", "пуш", "уведомлен",
        "геймификац", "социализац", "чат в мпк", "виджет", "сайт", "онлайн-запис",
        "клиентский опыт", "оплат через мпк", "оплат через приложен",
        "сторис", "стори", "видеокарточк", "удобств клиент", "интерфейс",
        "электронн сертификат", "электронн договор",
    ],
    "Интеграции и техническое": [
        "интеграц", "api", "1с", "1с фитнес", "битрикс", "amocrm", "amo crm",
        "yclients", "mobifitness", "mobi", "wallet", "wildberries", "склад",
        "касс atol", "atol", "эквайринг", "облачн касс", "телеграм-бот клуб",
        "интегриров", "ip-телефон", "перенос интеграц", "поддержка", "техподдерж",
        "регламент подключен", "ai-помощник", "gpt", "нейросет", "ии",
        "домен", "ссылк свой", "ссылк на свой", "внутренн ссылк", "pro версия",
    ],
    # 12-й блок — личная эффективность владельца (добавлен по Стасу из Stretch House)
    "Личная эффективность владельца": [
        "своим временем", "своё время", "режим сна", "режим питан", "режим работ",
        "тайм-менеджмент", "продуктивн", "энерги", "дисциплин",
        "стратегическое планир", "целеполаган", "приоритет", "цели на",
        "all hands", "корпоративн культур", "командный дух", "ценност компани",
        "вовлечённост сотрудник", "мотивац команд",
        "лично участвоват", "лично оценив", "лично собеседовани",
        "владелец", "лпр", "не быть узким местом лично",
    ],
}

# Расширения существующих блоков (добавляю чтобы 4 неклассифицированных попали)
BLOCKS["СКУД и доступ"].extend(["одновременн вход", "одновременн", "контролировать вход"])
BLOCKS["Клиенты и удержание"].extend(["корпоративн карт", "семейн карт", "карта для семь"])
BLOCKS["Запуск и переход"].extend(["электронн подписан", "электронн договор", "цифров подпис"])


def normalize(s):
    """Простая нормализация: lowercase, удалить ё→е, спецсимволы → пробелы"""
    if not s:
        return ""
    s = s.lower().replace("ё", "е")
    s = re.sub(r"[^а-яa-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Стоп-слова (мешают кластеризации, везде встречаются)
STOPWORDS = set("""
и в на с по для не как от из до за к у о а но или что чтобы
это того этого этом этой этих эти таких такого такие как
я мы он она они есть быть мочь могут можно надо нужно
если когда где который которое которые которых
свой свою свои свое все вся всё всем всеми вся всю
очень более менее много мало же ли бы под над через
""".split())


def stem_word(w):
    """Очень примитивный стеммер: убираем характерные окончания."""
    if len(w) <= 4:
        return w
    for suf in ["ования", "ование", "ованию", "ованиями", "ованиях",
                "ировать", "ирование", "ировал", "ируем",
                "ение", "ения", "ении", "ением", "ениях", "ениями",
                "ость", "ости", "остью", "остей",
                "ами", "ями", "ами", "ому", "ему", "ого", "его",
                "ыми", "ими", "ую", "юю", "ого", "его",
                "ться", "тся", "лся", "лась", "лись", "лось",
                "ать", "ять", "еть", "ить", "оть",
                "ие", "ий", "ия", "ыи", "ые", "ой", "ей",
                "ам", "ям", "ах", "ях", "ов", "ев",
                "а", "я", "о", "е", "у", "ю", "ы", "и"]:
        if w.endswith(suf) and len(w) - len(suf) >= 4:
            return w[:-len(suf)]
    return w


def tokens(s):
    """Возвращает список нормализованных «стемов»."""
    n = normalize(s)
    out = []
    for w in n.split():
        if w in STOPWORDS or len(w) < 3:
            continue
        out.append(stem_word(w))
    return out


def tokenized_text(s):
    return " ".join(tokens(s))


def score_block(text, keywords):
    """Сколько ключевых слов блока встречается в тексте."""
    n = normalize(text)
    return sum(1 for kw in keywords if kw in n)


def assign_block(row):
    """Назначить блок джобу по максимальному совпадению ключевых слов в want_short + so_that + want_result."""
    combined = " ".join([row.get(f, "") for f in ["want_short", "want_result", "so_that", "context_when"]])
    scores = {b: score_block(combined, kws) for b, kws in BLOCKS.items()}
    best_block = max(scores, key=scores.get)
    best_score = scores[best_block]
    if best_score == 0:
        return ("⚠️ Неклассифицировано", 0, scores)
    return (best_block, best_score, scores)


# Карта концептов: если в want_short встречается слово, оно даёт сильный сигнал на кластер.
# Внутри одного блока джобы с одинаковым CONCEPT_TAG будут объединены.
# Канонические названия кластеров (для красивого отображения вместо случайной формулировки из интервью)
CONCEPT_NAMES = {
    "зарплата_тренеров": "Считать зарплату тренеров без Excel и ошибок",
    "выручка_контроль": "Контролировать выручку и продажи в реальном времени",
    "реализация_vs_касса": "Видеть реальную прибыль, а не кассу (реализация vs касса)",
    "подписки_рекурренты": "Управлять подписками и рекуррентными списаниями",
    "абонементы_управление": "Создавать и управлять абонементами с разной номенклатурой",
    "оплаты_задолженност": "Отслеживать оплаты и задолженности клиентов",
    "акции_скидки": "Управлять акциями, скидками и промокодами",
    "tco_доплат": "Снизить TCO стека и избавиться от доплат за каждую фичу",
    "налоги_оферта": "Электронные договоры и оферта (юридически значимо)",
    "корпоративные_карты": "Корпоративные / семейные карты лояльности",
    "зарплата_сложная": "Считать сложную KPI-зарплату без ошибок",
    "найм_адаптация": "Нанимать и адаптировать сотрудников",
    "контроль_тренеров": "Контролировать качество работы тренеров",
    "переходы_тренеров": "Управлять переходами тренеров между клубами",
    "роли_доступ": "Управлять ролями и доступом сотрудников в системе",
    "адмы": "Управлять и контролировать администраторов на ресепшене",
    "защита_базы": "Защитить базу клиентов от увода тренерами",
    "удержание_отток": "Видеть кто «засыпает» и возвращать до отмены (удержание)",
    "рассылки_коммуникац": "Делать массовые рассылки и автоматические уведомления клиентам",
    "связь_родитель_ребёнок": "Связать аккаунты родителя и ребёнка",
    "программа_лояльности": "Программа лояльности с бонусами и «приведи друга»",
    "отзывы_обратная_связь": "Собирать отзывы и обратную связь клиентов",
    "клиентская_база": "Вести полную клиентскую базу и историю клиента",
    "коммуникация_мессенджеры": "Интегрировать все мессенджеры в одно окно (WhatsApp/TG/Max/VK/Instagram)",
    "виджеты_запись": "Виджеты онлайн-записи на сайт",
    "посещения_отметки": "Отмечать посещения клиентов (QR/чекин/ручная)",
    "расписание_создание": "Гибко составлять расписание занятий",
    "групповые_занят": "Управлять групповыми занятиями и записью",
    "персональные_занят": "Управлять персональными тренировками и записью",
    "аренда": "Управлять арендой залов / кортов",
    "лиды_воронка": "Лидогенерация и воронка продаж (заявки → клиент)",
    "контекст_таргет": "Контекстная реклама и таргет",
    "geo_маркетинг": "Локальный геомаркетинг (ЖК-чаты, штендеры, виндеры)",
    "smm_контент": "SMM и контент-маркетинг",
    "автономный_вход": "Автономный вход клиента в зал по QR/приложению",
    "кошкуд_дешевле": "Дешёвый СКУД для масштабирования (домофоны вместо Сигур/PocketKey)",
    "антифрод": "Антифрод и контроль несанкционированного прохода",
    "не_быть_узким": "Не быть узким местом — снять себя с операционки",
    "одно_окно": "Объединить всё в одном окне / единой системе",
    "масштабирование": "Масштабироваться (открывать новые точки / стать сетью)",
    "автоматизация": "Автоматизировать рутинные процессы",
    "перенос_базы": "Перенести базу клиентов без потерь при переходе на новую систему",
    "выбор_crm": "Выбрать CRM и перейти со старой системы",
    "запуск_бизнеса": "Запустить студию/клуб с нуля с готовыми процессами",
    "приложение_бренд": "Брендированное мобильное приложение под клуб",
    "приложение_удобство": "Удобный интерфейс мобильного приложения для клиента",
    "приложение_соц": "Социальные функции в МПК (чат, геймификация, знакомства)",
    "уведомления_клиент": "Push-уведомления клиентам",
    "интеграция_1с": "Интеграция с 1С (бухгалтерия)",
    "ai_gpt": "AI/GPT в операционных задачах владельца",
    "ip_телефония": "IP-телефония для работы с клиентами",
    "касса_atol": "Касса (Atol/облачная) и эквайринг",
    "склад_товар": "Учёт спортпита и товаров",
    "аналитика_бизнеса": "BI-аналитика и дашборды по бизнесу",
    "обзор_сети": "Обзор и сравнение точек сети",
    "тайм_менеджмент": "Управлять личным временем и энергией (режим, сон, спорт)",
    "стратегия": "Стратегическое планирование и целеполагание ЛПР",
    "корпкультура": "Поддерживать корпоративную культуру (All Hands, ценности)",
}


CONCEPT_KEYWORDS = {
    # Финансы
    "зарплата_тренеров": ["зарплат", "оплата труда", "процент тренер", "kpi тренер", "мотивац тренер"],
    "выручка_контроль": ["выручк", "доход", "финрез", "прибыл", "продаж и выручк"],
    "реализация_vs_касса": ["реализац", "касс", "дебиторк", "реальн прибыл", "реальн деньг"],
    "подписки_рекурренты": ["рекуррент", "подписк", "автопродлен", "автоматическ списан"],
    "абонементы_управление": ["абонемент", "учёт абонемент", "пакет услуг"],
    "оплаты_задолженност": ["задолженност", "долг клиент", "оплат и возврат"],
    "акции_скидки": ["акци", "скидк", "промокод", "сертификат"],
    "tco_доплат": ["tco", "доплат", "стоимост стека", "пакет тарифа"],
    "налоги_оферта": ["налог", "ндс", "оферт", "юридическ", "договор", "электронн подпис", "цифров подпис"],
    "корпоративные_карты": ["корпоративн карт", "семейн карт", "карта для семь"],
    # Персонал
    "зарплата_сложная": ["сложн зарплат", "зарплата по kpi", "зарплата по факту"],
    "найм_адаптация": ["найм", "обучен сотрудник", "адаптац", "стажир", "ввод в курс"],
    "контроль_тренеров": ["контролировать тренер", "проверять качеств тренировк", "запис тренировок"],
    "переходы_тренеров": ["переход тренер", "переман", "новый клуб тренер"],
    "роли_доступ": ["роль доступ", "ограничить доступ"],
    "адмы": ["админ", "ресепшен", "стойка ресепшен"],
    "защита_базы": ["защита баз", "уведут клиент", "nda тренер", "уход тренер с баз"],
    # Клиенты
    "удержание_отток": ["удержан", "отток", "засыпа", "ушедш", "вернуть до отмен"],
    "рассылки_коммуникац": ["рассылк", "массов рассылк", "уведомлен клиент", "напоминан"],
    "связь_родитель_ребёнок": ["родител ребенк", "родительск", "связ ребенк"],
    "программа_лояльности": ["лояльност", "бонус", "приведи друг"],
    "отзывы_обратная_связь": ["отзыв", "обратная связь", "кастдев"],
    "клиентская_база": ["база клиент", "учёт клиент", "карточк клиент", "анкет"],
    "коммуникация_мессенджеры": ["мессенджер", "whatsapp", "telegram", "telegram-бот", "max ", "instagram", "вконтакт"],
    # Расписание
    "виджеты_запись": ["виджет запис", "онлайн запис", "запис без регистрац", "запись в три клика"],
    "посещения_отметки": ["отмечать посещен", "отметить посещен", "посещаемост", "qr-чекин", "чекин"],
    "расписание_создание": ["составить расписан", "расписан занят", "копировать расписан", "повторяющ запис"],
    "групповые_занят": ["групповы заняти", "группов класс"],
    "персональные_занят": ["персональн", "запис на персональн"],
    "аренда": ["аренда залов", "аренда корт", "брониров слот"],
    # Маркетинг/коммуникация
    "лиды_воронка": ["лид", "воронк", "заявк из канал", "лидоген"],
    "контекст_таргет": ["контекстн реклам", "таргет", "ya direct"],
    "geo_маркетинг": ["ЖК-чат", "чат жк", "штендер", "виндер", "локальн маркетинг", "геомаркетинг"],
    "smm_контент": ["соцсет", "smm", "контент"],
    # СКУД
    "автономный_вход": ["автономн вход", "qr на вход", "вход по qr", "автоматическ вход"],
    "кошкуд_дешевле": ["домофон", "macwox", "сигур", "pocketkey", "стоимост скуд", "дешев скуд"],
    "антифрод": ["фрод", "одновременн вход", "несанкционирован", "видеоналит"],
    # Операционка
    "не_быть_узким": ["узкое место", "не быть узким", "снять с себя", "освободи себя", "минимизир ручн"],
    "одно_окно": ["в одном окне", "единая систем", "одно окно"],
    "масштабирование": ["масштабир", "стать сетью", "построить сеть"],
    "автоматизация": ["автоматизац процес", "автоматизирова процес"],
    # Запуск/переход
    "перенос_базы": ["перенос баз", "миграц баз", "перенос клиент"],
    "выбор_crm": ["выбор crm", "выбор систем", "переход на нов"],
    "запуск_бизнеса": ["запустить бизнес", "запустить студ", "запустить клуб", "открыть студ"],
    # МПК
    "приложение_бренд": ["брендир приложен", "приложен под бренд", "мобильн приложен брендир"],
    "приложение_удобство": ["удобство приложен", "интерфейс приложен", "мпк удобн"],
    "приложение_соц": ["социализац в мпк", "чат в мпк", "общение в приложен", "геймификац"],
    "уведомления_клиент": ["уведомлен клиент", "пуш"],
    # Интеграции
    "интеграция_1с": ["интеграц 1с", "1с буха"],
    "ai_gpt": ["gpt", "нейросет", "ai-помощник"],
    "ip_телефония": ["ip телефон", "телефон"],
    "касса_atol": ["atol", "касса atol", "облачн касс"],
    "склад_товар": ["спортпит", "склад товар"],
    # Аналитика
    "аналитика_бизнеса": ["аналитик", "дашборд", "отчёт", "выгрузить excel"],
    "обзор_сети": ["сравн точк", "обзор по сети"],
    # Личная эффективность
    "тайм_менеджмент": ["своим временем", "режим сна", "режим питан", "продуктивн"],
    "стратегия": ["стратегическ планирован", "целеполаган", "цели и задач", "приоритет на"],
    "корпкультура": ["all hands", "корпоративн культур", "командный дух", "ценност компани"],
}


def detect_concepts(text):
    """Какие концепты упоминаются в тексте."""
    n = normalize(text)
    found = set()
    for tag, kws in CONCEPT_KEYWORDS.items():
        for kw in kws:
            if kw in n:
                found.add(tag)
                break
    return found


def cluster_within_block(rows, threshold=0.30):
    """1) Pre-cluster по концептам, 2) TF-IDF refinement для остатков."""
    if len(rows) <= 1:
        return [[0]] if rows else []

    # Шаг 1: каждому row найти его «доминантный концепт» (первый из найденных)
    concept_groups = defaultdict(list)
    no_concept = []
    for i, r in enumerate(rows):
        text = r["want_short"] + " " + r.get("so_that", "")
        cs = detect_concepts(text)
        if cs:
            best_tag = sorted(cs)[0]
            concept_groups[best_tag].append(i)
        else:
            no_concept.append(i)
    # Запомнить tag → cluster_idx (для последующего именования)
    cluster_within_block._last_tags = list(concept_groups.keys())

    # Шаг 2: на остатках без концепта — TF-IDF кластеризация
    extra_groups = []
    if len(no_concept) > 1:
        texts = [tokenized_text(rows[i]["want_short"] + " " + rows[i].get("so_that", "")) for i in no_concept]
        if any(texts):
            try:
                vec = TfidfVectorizer(analyzer="word", ngram_range=(1, 2), min_df=1)
                X = vec.fit_transform(texts)
                sim = cosine_similarity(X)
                # union-find
                parent = list(range(len(no_concept)))
                def find(x):
                    while parent[x] != x:
                        parent[x] = parent[parent[x]]
                        x = parent[x]
                    return x
                for a in range(len(no_concept)):
                    for b in range(a + 1, len(no_concept)):
                        if sim[a, b] >= threshold:
                            ra, rb = find(a), find(b)
                            if ra != rb:
                                parent[ra] = rb
                groups = defaultdict(list)
                for idx in range(len(no_concept)):
                    groups[find(idx)].append(no_concept[idx])
                extra_groups = list(groups.values())
            except Exception:
                extra_groups = [[i] for i in no_concept]
        else:
            extra_groups = [[i] for i in no_concept]
    elif len(no_concept) == 1:
        extra_groups = [no_concept]

    # Объединить
    return list(concept_groups.values()) + extra_groups


def _legacy_cluster_within_block(rows, threshold=0.30):
    """Старая версия — оставлена для справки."""
    if len(rows) <= 1:
        return [[0]] if rows else []
    texts = [tokenized_text(r["want_short"] + " " + r.get("so_that", "")) for r in rows]
    vec = TfidfVectorizer(analyzer="word", ngram_range=(1, 2), min_df=1)
    try:
        X = vec.fit_transform(texts)
    except Exception:
        return [[i] for i in range(len(rows))]
    sim = cosine_similarity(X)

    # Простой агломеративный greedy: для каждой пары с sim >= threshold — мерж
    parent = list(range(len(rows)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    n = len(rows)
    for i in range(n):
        for j in range(i + 1, n):
            if sim[i, j] >= threshold:
                union(i, j)

    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    return list(groups.values())


def aggregate_cluster(rows_in_cluster):
    """Объединить джобы в кластере → одна запись с frequency/segments/sources."""
    # Канонический want_short: наибольшая длина (полнее формулировка)
    canonical = max(rows_in_cluster, key=lambda r: len(r["want_short"]))

    # Уникальные значения по полям
    sources = sorted(set(r["interview_label"] for r in rows_in_cluster if r["interview_label"]))
    seg_b = sorted(set(r["business_segment"] for r in rows_in_cluster if r["business_segment"] and r["business_segment"] != "?"))
    seg_j = sorted(set(r["job_segment"] for r in rows_in_cluster if r["job_segment"] and r["job_segment"] != "?"))
    sizes = Counter(r["job_size"] for r in rows_in_cluster if r["job_size"])
    dominant_size = sizes.most_common(1)[0][0] if sizes else ""

    # Все want_short — для проверки человеком (если кластер собрал что-то странное)
    all_wants = " | ".join(sorted(set(r["want_short"] for r in rows_in_cluster)))

    # so_that: берём самый длинный непустой
    so_thats = sorted(set(r.get("so_that", "") for r in rows_in_cluster if r.get("so_that")), key=len, reverse=True)
    so_that = so_thats[0] if so_thats else ""

    # context_when: самый длинный
    contexts = sorted(set(r.get("context_when", "") for r in rows_in_cluster if r.get("context_when")), key=len, reverse=True)
    context = contexts[0] if contexts else ""

    # Цитаты — все непустые через перевод строки
    quotes = sorted(set(r.get("lpr_quote", "") for r in rows_in_cluster if r.get("lpr_quote")), key=len, reverse=True)
    quote = " || ".join(quotes[:3])  # топ-3 самые длинные

    # Барьеры / драйверы / fitbase_value — самые длинные
    def longest(field):
        vals = sorted(set(r.get(field, "") for r in rows_in_cluster if r.get(field)), key=len, reverse=True)
        return vals[0] if vals else ""

    return {
        "job_size": dominant_size,
        "want_short": canonical["want_short"],
        "context_when": context,
        "so_that": so_that,
        "fitbase_value": longest("fitbase_value"),
        "barriers": longest("barriers"),
        "drivers": longest("drivers"),
        "previous_solution_problems": longest("previous_solution_problems"),
        "lpr_quote": quote,
        "frequency": len(rows_in_cluster),
        "business_segments": ", ".join(seg_b),
        "job_segments": ", ".join(seg_j),
        "sources_count": len(sources),
        "sources": "; ".join(sources),
        "all_variants": all_wants,
    }


def main():
    with open(IN_CSV, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    # 1. Назначить блок каждой строке
    by_block = defaultdict(list)
    unclassified = []
    # Ручные override для оставшихся пограничных случаев
    MANUAL_OVERRIDES = {
        "электронное подписание документов": "Финансы и учёт",
        "корпоративные карты для семей": "Клиенты и удержание",
    }
    for r in rows:
        if not r["want_short"] or r["want_short"].startswith("⚠️"):
            continue
        # Проверка ручных override
        ws_low = normalize(r["want_short"])
        forced_block = None
        for key, blk in MANUAL_OVERRIDES.items():
            if key in ws_low:
                forced_block = blk
                break
        if forced_block:
            by_block[forced_block].append(r)
            continue
        block, score, _ = assign_block(r)
        by_block[block].append(r)
        if block == "⚠️ Неклассифицировано":
            unclassified.append(r)

    # 2. Внутри каждого блока — кластеризовать с привязкой к концептам
    clustered = []
    for block, block_rows in by_block.items():
        # Концептные группы
        concept_groups = defaultdict(list)
        no_concept = []
        for i, r in enumerate(block_rows):
            text = r["want_short"] + " " + r.get("so_that", "")
            cs = detect_concepts(text)
            if cs:
                best_tag = sorted(cs)[0]
                concept_groups[best_tag].append(i)
            else:
                no_concept.append(i)

        grp_idx = 0
        # Сначала концептные группы (с каноническим именем)
        for tag, indices in concept_groups.items():
            grp_idx += 1
            cluster_rows = [block_rows[i] for i in indices]
            agg = aggregate_cluster(cluster_rows)
            agg["block"] = block
            agg["cluster_id"] = f"{block[:3]}-{grp_idx:02d}"
            # Заменить want_short на каноническое имя концепта
            if tag in CONCEPT_NAMES:
                agg["canonical_name"] = CONCEPT_NAMES[tag]
            else:
                agg["canonical_name"] = agg["want_short"]
            agg["concept_tag"] = tag
            clustered.append(agg)

        # Затем no_concept — TF-IDF refinement
        if no_concept:
            extra_groups = []
            if len(no_concept) > 1:
                texts = [tokenized_text(block_rows[i]["want_short"] + " " + block_rows[i].get("so_that", "")) for i in no_concept]
                if any(texts):
                    try:
                        vec = TfidfVectorizer(analyzer="word", ngram_range=(1, 2), min_df=1)
                        X = vec.fit_transform(texts)
                        sim = cosine_similarity(X)
                        parent = list(range(len(no_concept)))
                        def find(x):
                            while parent[x] != x:
                                parent[x] = parent[parent[x]]
                                x = parent[x]
                            return x
                        for a in range(len(no_concept)):
                            for b in range(a + 1, len(no_concept)):
                                if sim[a, b] >= 0.30:
                                    ra, rb = find(a), find(b)
                                    if ra != rb:
                                        parent[ra] = rb
                        groups_dict = defaultdict(list)
                        for idx in range(len(no_concept)):
                            groups_dict[find(idx)].append(no_concept[idx])
                        extra_groups = list(groups_dict.values())
                    except Exception:
                        extra_groups = [[i] for i in no_concept]
                else:
                    extra_groups = [[i] for i in no_concept]
            else:
                extra_groups = [no_concept]

            for indices in extra_groups:
                grp_idx += 1
                cluster_rows = [block_rows[i] for i in indices]
                agg = aggregate_cluster(cluster_rows)
                agg["block"] = block
                agg["cluster_id"] = f"{block[:3]}-{grp_idx:02d}"
                agg["canonical_name"] = agg["want_short"]
                agg["concept_tag"] = ""
                clustered.append(agg)

    # Сортировка: по блоку, потом по frequency desc
    block_order = list(BLOCKS.keys()) + ["⚠️ Неклассифицировано"]
    clustered.sort(key=lambda x: (block_order.index(x["block"]) if x["block"] in block_order else 99, -x["frequency"]))

    # 3. Сохранить
    OUT_COLS = [
        "block", "job_size", "canonical_name",
        "frequency", "business_segments", "job_segments", "sources_count",
        "so_that", "context_when", "fitbase_value",
        "barriers", "drivers", "previous_solution_problems",
        "lpr_quote", "sources", "all_variants",
        "cluster_id", "concept_tag",
    ]
    out_csv = BASE / "jobs_clustered.csv"
    with open(out_csv, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=OUT_COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(clustered)
    print(f"✅ CSV: {out_csv}")
    print(f"📊 Было: {len(rows)} | Стало кластеров: {len(clustered)}")

    # 4. Отчёт
    report = BASE / "clustering_report.md"
    with open(report, "w", encoding="utf-8") as f:
        f.write("# Отчёт по кластеризации джобов\n\n")
        f.write(f"**Исходно:** {len(rows)} джобов  \n")
        f.write(f"**После кластеризации:** {len(clustered)} уникальных работ  \n\n")
        f.write("## Распределение по блокам\n\n")
        f.write("| Блок | Кол-во кластеров | Сырых джобов |\n|------|------|------|\n")
        for block in block_order:
            n_clusters = sum(1 for c in clustered if c["block"] == block)
            n_raw = sum(c["frequency"] for c in clustered if c["block"] == block)
            if n_clusters > 0:
                f.write(f"| {block} | {n_clusters} | {n_raw} |\n")

        if unclassified:
            f.write(f"\n## ⚠️ Не классифицировано ({len(unclassified)} строк)\n\n")
            f.write("Эти джобы не попали ни в один блок — нужно либо добавить ключевые слова, либо новый блок:\n\n")
            for r in unclassified[:30]:
                f.write(f"- **{r['want_short']}** _(из «{r['interview_label']}»)_\n")
            if len(unclassified) > 30:
                f.write(f"\n*(+ ещё {len(unclassified)-30} строк)*\n")

        f.write("\n## Топ-20 самых частотных кластеров\n\n")
        f.write("| # | Размер | Работа | Кол-во интервью | Сегменты |\n|---|---|---|---|---|\n")
        sorted_freq = sorted(clustered, key=lambda x: -x["frequency"])
        for i, c in enumerate(sorted_freq[:20], 1):
            name = c.get("canonical_name") or c["want_short"]
            f.write(f"| {i} | {c['job_size']} | {name[:90]} | **{c['frequency']}** | {c['business_segments']} |\n")

    print(f"✅ Отчёт: {report}")
    return clustered, unclassified


if __name__ == "__main__":
    main()
