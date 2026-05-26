#!/usr/bin/env python3
"""
Сборщик единой таблицы джобов Fitbase из всех источников.

Источники:
1. 19 CSV-файлов в Interview_Analysis/segments/ — основные данные с привязкой к интервью
2. jobs_and_pains/jobs.csv — 30 агрегированных джобов с УТП (без привязки к ЛПР)
3. MD JTBD-файлы по интервью:
   - analysis_catalina_jtbd.md (Каталина, эталон)
   - Knowledge/Interviews/Padel/denis_padel_jtbd.md (Денис, падел)
   - Knowledge/Interviews/sales_calls_1c_switch/0[1-5]_*.md (5 интервью переходов с 1С)
4. UrbanFit и Kochkin/SpeakerA Analysis — narrative-формат, не парсятся, добавляются как
   ссылка-плейсхолдер с notes="требует ручной обработки"

ВЫВОД:
- jobs_master.csv
- jobs_master.xlsx (через xlsx skill отдельно)
- jobs_master.html (интерактивный)
- verification_report.md
"""

import csv
import os
import re
from pathlib import Path
from collections import defaultdict

BASE_INTERVIEW = Path("/Users/romansemenov/Desktop/Projects/Work/Interview_Analysis")
BASE_KNOWLEDGE = Path("/Users/romansemenov/Desktop/Cloud Cod/Knowledge")
BASE_ROOT = Path("/Users/romansemenov/Desktop/Cloud Cod")
OUT = Path("/Users/romansemenov/Desktop/Cloud Cod/Knowledge/jobs_master_table")

SEGMENT_HINTS = {
    "network/dmitry_network_jtbd": {"S": "S3", "J": "J4", "label": "Дмитрий — сеть, JTBD-вариант"},
    "network/dmitry_network_no_admins": {"S": "S3", "J": "J4", "label": "Дмитрий — сеть без админов"},
    "network/dmitry_network_transcript": {"S": "S3", "J": "J4", "label": "Дмитрий — сеть, транскрипт"},
    "opening/opening_studio": {"S": "S1", "J": "J3", "label": "Ярослава — открытие студии пилатеса"},
    "switch_other_systems/oksana_two_studios_appevent": {"S": "S1", "J": "J1", "label": "Оксана — две детских студии (AppEvent)"},
    "switch_other_systems/switch_club_from_1c": {"S": "S2", "J": "J2", "label": "Клуб — переход с 1С"},
    "switch_other_systems/switch_club_from_sheet": {"S": "S2", "J": "J2", "label": "Клуб — переход с таблиц"},
    "switch_other_systems/switch_club_generic": {"S": "S2", "J": "J2", "label": "Клуб — общий switch-кейс"},
    "switch_other_systems/switch_from_1c": {"S": "S1", "J": "J1", "label": "Анна — танцы, переход с 1С"},
    "switch_other_systems/switch_from_1c_variant_2": {"S": "S1", "J": "J1", "label": "Студия — переход с 1С (вариант 2)"},
    "switch_other_systems/switch_from_bitrix": {"S": "S1", "J": "J1", "label": "Студия — переход с Bitrix"},
    "switch_other_systems/switch_mobi_vs_fitbase_manager": {"S": "S1", "J": "J2", "label": "Менеджер — Mobifitness vs Fitbase"},
    "switch_other_systems/switch_mobifitness": {"S": "S1", "J": "J2", "label": "Переход с Mobifitness"},
    "switch_other_systems/switch_rasul_dance_studio": {"S": "S1", "J": "J1", "label": "Расул — танцевальная студия"},
    "switch_other_systems/switch_sheet_amo_sosnovy_bor": {"S": "S2", "J": "J2", "label": "Сосновый Бор — клуб (Excel+AmoCRM)"},
    "switch_other_systems/switch_zenfitness_kraft_yc": {"S": "S2", "J": "J2", "label": "Zenfitness/Kraft — переход с YClients"},
    "switch_tables/switch_tables_lime_fitness": {"S": "S1", "J": "J1", "label": "Lime Fitness — хореография"},
    "switch_tables/switch_tables_studio_mobi_competitor": {"S": "S1", "J": "J1", "label": "Танцевальная студия — Mobi-конкурент"},
    "unmanned/speaker_a_unmanned": {"S": "S5", "J": "J1", "label": "Speaker A — зал без персонала"},
    "unmanned/speaker_a_unmanned_full": {"S": "S5", "J": "J1", "label": "Speaker A — зал без персонала (расширенный, 8 джобов)"},
    "unmanned/kochkin_unmanned": {"S": "S5", "J": "J3", "label": "Алексей Кочкин — зал без персонала (Екатеринбург, открытие)"},
    "network/stretch_house_stas": {"S": "S4", "J": "J4", "label": "Стас — Stretch House (франшиза, 65 студий)"},
    "network/urbanfit_full": {"S": "S3", "J": "J4", "label": "Urban Fit — сеть рекуррентов (Жан + Евгений, 12 клубов СПб)"},
}

# Финальный порядок: Размер → Сегмент → Хочу (название) → контекст и далее.
# Дубликаты НЕ попадают в публичный вывод — оставляем только canonical.
TARGET_COLUMNS = [
    "job_size",                    # 1. Размер
    "business_segment",             # 2. Бизнес-сегмент S1–S6
    "job_segment",                  # 3. Джоб-сегмент J1–J6
    "want_short",                   # 4. Название джоба / Хочу кратко
    "context_when",                 # 5. Когда (контекст, триггер)
    "want_result",                  # 6. Хочу получить
    "so_that",                      # 7. Чтобы
    "importance",                   # 8. Важность
    "satisfaction",                 # 9. Удовлетворённость
    "gap",                          # 10. Gap
    "fitbase_value",                # 11. Ценность Fitbase
    "previous_solution_problems",   # 12. Проблемы прошлого решения
    "drivers",                      # 13. Драйверы
    "barriers",                     # 14. Барьеры
    "lpr_quote",                    # 15. Цитата ЛПР
    "interview_label",              # 16. Интервью / ЛПР
    "business_profile",             # 17. Профиль бизнеса
    "lpr_profile",                  # 18. Профиль ЛПР
    "notes",                        # 19. Заметки
    "id",                           # 20. ID строки
]

# Внутренние поля (используются для группировки дубликатов, в вывод не попадают)
INTERNAL_FIELDS = ["source_file", "source_type", "duplicate_group", "is_canonical"]


def normalize_size(s):
    """«Big Job» → «Big», «Middle Job 1» → «Middle», «Small Job 5» → «Small»."""
    if not s:
        return ""
    s = s.strip()
    low = s.lower()
    if low.startswith("big"):
        return "Big"
    if low.startswith("middle"):
        return "Middle"
    if low.startswith("small"):
        return "Small"
    return s

HEADER_MAP = {
    "размер работы": "job_size",
    "уровень работы": "job_size",
    "короткое и запоминающее описание работы (=хочу)": "want_short",
    "короткое и запоминающееся описание работы (=хочу)": "want_short",
    "короткое описание работы (=хочу)": "want_short",
    "короткое и запоминающееся описание работы (хочу)": "want_short",
    "хочу (короткое описание)": "want_short",
    "когда (контекст, триггеры)": "context_when",
    "когда (контекст, особенности)": "context_when",
    "когда (контекст, психологика, опыт, триггер)": "context_when",
    "когда (контекст, психологические особенности, прошлый опыт, триггер)": "context_when",
    "когда (контекст)": "context_when",
    "хочу получить (ожидаемый результат)": "want_result",
    "хочу получить ожидаемый результат": "want_result",
    "хочу получить": "want_result",
    "чтобы получить (результат работы выше уровнем, эмоциональный эффект)": "so_that",
    "чтобы получить (результат выше уровнем, эмоциональный эффект)": "so_that",
    "чтобы получить ожидаемый результат работы выше уровнем и чувствовать себя по-другому": "so_that",
    "чтобы получить (результат работы выше уровнем)": "so_that",
    "чтобы (цель, эффект)": "so_that",
    "чтобы (цель)": "so_that",
    "важность выполнения работы": "importance",
    "важность выполнения работы (почему важно)": "importance",
    "ценность от фитбейс": "fitbase_value",
    "ценность от fitbase": "fitbase_value",
    "ценность от fitbase (если есть)": "fitbase_value",
    "ценность от fitbase (гипотеза)": "fitbase_value",
    "ценность, которую хотят получить от fitbase": "fitbase_value",
    "проблемы с предыдущим решением": "previous_solution_problems",
    "проблемы с предыдущим решением (1с фитнес и др.)": "previous_solution_problems",
    "проблемы с предыдущим решением (1с)": "previous_solution_problems",
    "проблемы с предыдущим/текущим решением (подробно)": "previous_solution_problems",
    "драйверы к выбору фитбейс": "drivers",
    "драйверы к выбору fitbase": "drivers",
    "драйверы к переходу/выбору решения": "drivers",
    "барьеры к выбору фитбейс": "barriers",
    "барьеры к выбору fitbase": "barriers",
    "цитаты лпр": "lpr_quote",
    "цитаты лпр (анна)": "lpr_quote",
    "цитаты лпр (собственника, если есть)": "lpr_quote",
    "цитаты лпр (если есть)": "lpr_quote",
    "про текущую программу": "notes",
}


def empty_record():
    # Создаём словарь со ВСЕМИ полями (включая внутренние), чтобы парсеры могли их использовать.
    # При записи в CSV/Excel будут использоваться только TARGET_COLUMNS.
    return {k: "" for k in TARGET_COLUMNS + INTERNAL_FIELDS}


def normalize_header(h):
    return (h or "").strip().lower()


def find_header_row(rows):
    for i, r in enumerate(rows):
        if not r:
            continue
        first = (r[0] or "").strip().lower()
        if first in ("размер работы", "уровень работы"):
            return i
    return None


def extract_metadata(rows, header_idx):
    business = []
    lpr = []
    for i in range(header_idx):
        r = rows[i]
        if not r or len(r) < 2:
            continue
        key = (r[0] or "").strip().lower()
        val = (r[1] or "").strip()
        if not val:
            continue
        if "лпр" in key or "лицо, принимающ" in key:
            lpr.append(val)
        elif "бизнес" in key or "особенност" in key or "тип бизнес" in key or "сегмент" in key or "предыдущ" in key:
            business.append(f"{r[0]}: {val}" if r[0] else val)
        elif "информация о бизнесе" in key:
            business.append(val)
    return " | ".join(business), " | ".join(lpr)


def parse_segment_csv(path: Path):
    rel = path.relative_to(BASE_INTERVIEW).as_posix()
    segment_key = "/".join(rel.split("/")[1:3])
    hint = SEGMENT_HINTS.get(segment_key, {})

    with open(path, encoding="utf-8") as f:
        rows = list(csv.reader(f))

    header_idx = find_header_row(rows)
    if header_idx is None:
        for i, r in enumerate(rows):
            if r and len(r) >= 2 and "уровень работы" in (r[0] or "").lower():
                header_idx = i
                break

    if header_idx is None:
        rec = empty_record()
        rec.update({
            "source_file": rel,
            "source_type": "csv_segment",
            "interview_label": hint.get("label", segment_key),
            "business_segment": hint.get("S", ""),
            "job_segment": hint.get("J", ""),
            "notes": "ПАРСЕР НЕ НАШЁЛ ЗАГОЛОВОК — проверить вручную",
        })
        return [rec]

    business, lpr = extract_metadata(rows, header_idx)
    headers = rows[header_idx]
    col_map = {}
    for i, h in enumerate(headers):
        nk = normalize_header(h)
        if nk in HEADER_MAP:
            col_map[i] = HEADER_MAP[nk]

    out = []
    for r in rows[header_idx + 1:]:
        if not r or not any(cell.strip() for cell in r if cell):
            continue
        if not r[0].strip():
            continue
        if normalize_header(r[0]) in HEADER_MAP and HEADER_MAP[normalize_header(r[0])] == "job_size":
            continue
        rec = empty_record()
        rec["source_file"] = rel
        rec["source_type"] = "csv_segment"
        rec["interview_label"] = hint.get("label", segment_key)
        rec["business_segment"] = hint.get("S", "")
        rec["job_segment"] = hint.get("J", "")
        rec["business_profile"] = business
        rec["lpr_profile"] = lpr
        for i, key in col_map.items():
            if i < len(r):
                val = (r[i] or "").strip()
                if rec[key] and val:
                    rec[key] = rec[key] + " | " + val
                elif val:
                    rec[key] = val
        out.append(rec)
    return out


def parse_aggregated_csv(path: Path):
    rel = path.relative_to(BASE_INTERVIEW).as_posix()
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        out = []
        for row in reader:
            rec = empty_record()
            rec["source_file"] = rel
            rec["source_type"] = "csv_aggregated"
            rec["interview_label"] = "АГРЕГИРОВАНО (без привязки к конкретному интервью)"
            rec["job_size"] = row.get("Размер джобы", "").strip()
            rec["want_short"] = row.get("Короткое описание (=хочу)", "").strip()
            rec["context_when"] = row.get("Когда", "").strip()
            rec["want_result"] = row.get("Хочу", "").strip()
            rec["so_that"] = row.get("Чтобы", "").strip()
            rec["importance"] = row.get("Важность", "").strip()
            rec["fitbase_value"] = row.get("Фича от Fitbase", "").strip()
            rec["previous_solution_problems"] = row.get("Проблемы прошлого решения", "").strip()
            rec["drivers"] = row.get("Драйверы", "").strip()
            rec["barriers"] = row.get("Барьеры", "").strip()
            utp = row.get("УТП Fitbase (фичи + рекламная подача)", "").strip()
            freq = row.get("Частотность", "").strip()
            rec["notes"] = (f"Частотность: {freq}" if freq else "") + (f" | УТП: {utp}" if utp else "")
            if rec["want_short"] or rec["want_result"]:
                out.append(rec)
    return out


# ==== ПАРСЕРЫ MD ====

def parse_catalina(path: Path):
    """analysis_catalina_jtbd.md — структура:
    ### BIG JOB
    **Title**

    ### MIDDLE JOBS
    *   **MJ-1: Title.** Description
    *   **MJ-2: Title.** Description

    ### SMALL JOBS
    *   text
    *   text
    """
    rel = str(path.relative_to(BASE_ROOT))
    text = path.read_text(encoding="utf-8")
    jobs = []

    # BIG JOB
    m = re.search(r"###\s+BIG JOB\s*\n+\*\*([^*]+)\*\*", text)
    if m:
        rec = empty_record()
        rec.update({
            "source_file": rel, "source_type": "md_jtbd",
            "interview_label": "Каталина Алексеева — пилатес/реабилитация, Гатчина (эталон AJTBD)",
            "business_segment": "S1", "job_segment": "J2",
            "job_size": "Big",
            "want_short": m.group(1).strip().rstrip("."),
        })
        jobs.append(rec)

    # MIDDLE JOBS
    middle_section = re.search(r"###\s+MIDDLE JOBS\s*\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL)
    if middle_section:
        block = middle_section.group(1)
        for m in re.finditer(r"^\*\s+\*\*MJ-\d+:\s+([^*]+?)\*\*\s*(.*)$", block, re.MULTILINE):
            title = m.group(1).strip().rstrip(".")
            desc = m.group(2).strip()
            rec = empty_record()
            rec.update({
                "source_file": rel, "source_type": "md_jtbd",
                "interview_label": "Каталина Алексеева — пилатес/реабилитация, Гатчина",
                "business_segment": "S1", "job_segment": "J2",
                "job_size": "Middle",
                "want_short": title,
                "context_when": desc,
            })
            jobs.append(rec)

    # SMALL JOBS
    small_section = re.search(r"###\s+SMALL JOBS\s*\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL)
    if small_section:
        block = small_section.group(1)
        for m in re.finditer(r"^\*\s+([^\n*][^\n]+)$", block, re.MULTILINE):
            text_line = m.group(1).strip().rstrip(".")
            if not text_line:
                continue
            rec = empty_record()
            rec.update({
                "source_file": rel, "source_type": "md_jtbd",
                "interview_label": "Каталина Алексеева — пилатес/реабилитация, Гатчина",
                "business_segment": "S1", "job_segment": "J2",
                "job_size": "Small",
                "want_short": text_line,
            })
            jobs.append(rec)

    return jobs


def parse_padel(path: Path):
    """denis_padel_jtbd.md — Big + 2 Middle + 7 Small из markdown-таблицы."""
    rel = str(path.relative_to(BASE_ROOT))
    text = path.read_text(encoding="utf-8")
    jobs = []

    LABEL = "Денис — падл-клуб (старт)"
    S, J = "S6", "J4"

    # Big Job
    big = re.search(r"##\s+Big Job\s*\n+\*\*([^*]+)\*\*\s*\n(.*?)(?=\n##\s|\Z)", text, re.DOTALL)
    if big:
        rec = empty_record()
        block = big.group(2)
        rec.update({
            "source_file": rel, "source_type": "md_jtbd",
            "interview_label": LABEL, "business_segment": S, "job_segment": J,
            "job_size": "Big", "want_short": big.group(1).strip(),
        })
        for f, pat in [
            ("context_when", r"\*\*Когда:\*\*\s*([^\n]+)"),
            ("want_result", r"\*\*Хочу:\*\*\s*([^\n]+)"),
            ("so_that", r"\*\*Чтобы:\*\*\s*([^\n]+)"),
            ("fitbase_value", r"\*\*Ценность Fitbase:\*\*\s*([^\n]+)"),
            ("barriers", r"\*\*Барьер[ыы]?:\*\*\s*([^\n]+)"),
        ]:
            m2 = re.search(pat, block)
            if m2:
                rec[f] = m2.group(1).strip()
        q = re.search(r">\s+«([^»]+)»", block)
        if q:
            rec["lpr_quote"] = q.group(1).strip()
        jobs.append(rec)

    # Middle Jobs — ### MJ-N. Title
    for m in re.finditer(r"###\s+MJ-(\d+)\.\s+([^\n]+)\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL):
        title = m.group(2).strip()
        block = m.group(3)
        rec = empty_record()
        rec.update({
            "source_file": rel, "source_type": "md_jtbd",
            "interview_label": LABEL, "business_segment": S, "job_segment": J,
            "job_size": "Middle", "want_short": title,
        })
        for f, pat in [
            ("context_when", r"\*\*Когда:\*\*\s*([^\n]+)"),
            ("want_result", r"\*\*Хочу:\*\*\s*([^\n]+)"),
            ("so_that", r"\*\*Чтобы:\*\*\s*([^\n]+)"),
            ("fitbase_value", r"\*\*Ценность Fitbase:\*\*\s*([^\n]+)"),
            ("barriers", r"\*\*Барьер[ыы]?:\*\*\s*([^\n]+)"),
        ]:
            m2 = re.search(pat, block)
            if m2:
                rec[f] = m2.group(1).strip()
        q = re.search(r">\s+«([^»]+)»", block)
        if q:
            rec["lpr_quote"] = q.group(1).strip()
        jobs.append(rec)

    # Small Jobs из markdown-таблицы
    small_section = re.search(r"##\s+Small Jobs[^\n]*\n+(.*?)(?=\n##\s|\Z)", text, re.DOTALL)
    if small_section:
        for line in small_section.group(1).split("\n"):
            line = line.strip()
            if not line.startswith("|"):
                continue
            cells = [c.strip() for c in line.split("|")[1:-1]]  # remove first/last empty
            if len(cells) < 2:
                continue
            # пропустить заголовок и разделитель
            if cells[0] == "#" or set(cells[0]) <= set("-:"):
                continue
            rec = empty_record()
            rec.update({
                "source_file": rel, "source_type": "md_jtbd",
                "interview_label": LABEL, "business_segment": S, "job_segment": J,
                "job_size": "Small", "want_short": cells[1] if len(cells) > 1 else "",
                "lpr_quote": cells[2] if len(cells) > 2 else "",
            })
            jobs.append(rec)

    return jobs


def parse_sales_call_1c(path: Path, business_seg, job_seg, label):
    """Файлы 0[1-5]_*.md формата Левита: ### BIG JOB / ### MIDDLE JOBS / **MJ-N: Title** / ### SMALL JOBS."""
    rel = str(path.relative_to(BASE_ROOT))
    text = path.read_text(encoding="utf-8")
    jobs = []

    # BIG JOB: ### BIG JOB\n**Title**
    big = re.search(r"###\s+BIG JOB\s*\n+\*\*([^*]+)\*\*\s*\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL)
    if big:
        rec = empty_record()
        block = big.group(2)
        rec.update({
            "source_file": rel, "source_type": "md_jtbd",
            "interview_label": label, "business_segment": business_seg, "job_segment": job_seg,
            "job_size": "Big", "want_short": big.group(1).strip(),
        })
        for f, pat in [
            ("context_when", r"\*\*Когда:\*\*\s*([^\n]+)"),
            ("want_result", r"\*\*Хочу:\*\*\s*([^\n]+)"),
            ("so_that", r"\*\*Чтобы:\*\*\s*([^\n]+)"),
        ]:
            m2 = re.search(pat, block)
            if m2:
                rec[f] = m2.group(1).strip()
        jobs.append(rec)

    # MIDDLE JOBS: **MJ-N: Title**\n- *Контекст:*/*Хочу:*/*Чтобы:*/*Драйверы:*/*Барьеры:*
    middle_section = re.search(r"###\s+MIDDLE JOBS\s*\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL)
    if middle_section:
        block = middle_section.group(1)
        # Каждый MJ блок
        for m in re.finditer(r"\*\*MJ-\d+:\s+([^*]+?)\*\*\s*\n(.*?)(?=\n\*\*MJ-\d+:|\Z)", block, re.DOTALL):
            title = m.group(1).strip()
            sub = m.group(2)
            rec = empty_record()
            rec.update({
                "source_file": rel, "source_type": "md_jtbd",
                "interview_label": label, "business_segment": business_seg, "job_segment": job_seg,
                "job_size": "Middle", "want_short": title,
            })
            for f, pat in [
                ("context_when", r"-\s*\*Контекст:\*\s*([^\n]+)"),
                ("want_result", r"-\s*\*Хочу:\*\s*([^\n]+)"),
                ("so_that", r"-\s*\*Чтобы:\*\s*([^\n]+)"),
                ("importance", r"-\s*\*Важность[^:]*:\*\s*([^\n]+)"),
                ("satisfaction", r"-\s*\*Удовлетвор[её]нность[^:]*:\*\s*([^\n]+)"),
                ("previous_solution_problems", r"-\s*\*Проблемы[^:]*:\*\s*([^\n]+)"),
                ("drivers", r"-\s*\*Драйверы[^:]*:\*\s*([^\n]+)"),
                ("barriers", r"-\s*\*Барьер[ыы]?[^:]*:\*\s*([^\n]+)"),
            ]:
                m2 = re.search(pat, sub)
                if m2:
                    rec[f] = m2.group(1).strip()
            jobs.append(rec)

    # SMALL JOBS: bullet list
    small_section = re.search(r"###\s+SMALL JOBS\s*\n(.*?)(?=\n###\s|\n##\s|\Z)", text, re.DOTALL)
    if small_section:
        for line in small_section.group(1).split("\n"):
            line = line.strip()
            # bullet items: "- text" or "* text"
            if not (line.startswith("- ") or line.startswith("* ")):
                continue
            content = line[2:].strip()
            # очистить markdown markup
            content = re.sub(r"\*\*([^*]+)\*\*", r"\1", content)
            if not content or len(content) < 5:
                continue
            rec = empty_record()
            rec.update({
                "source_file": rel, "source_type": "md_jtbd",
                "interview_label": label, "business_segment": business_seg, "job_segment": job_seg,
                "job_size": "Small", "want_short": content,
            })
            jobs.append(rec)

    return jobs


def add_unparseable_reference(path, business_seg, job_seg, label, note):
    """Для narrative-MD добавляем одну плейсхолдер-строку с указанием на необходимость ручной обработки."""
    rel = str(path.relative_to(BASE_ROOT))
    rec = empty_record()
    rec.update({
        "source_file": rel, "source_type": "md_jtbd_narrative",
        "interview_label": label, "business_segment": business_seg, "job_segment": job_seg,
        "job_size": "",
        "want_short": f"⚠️ Narrative-формат, не парсится — см. {path.name}",
        "notes": note,
    })
    return [rec]


def main():
    all_records = []
    csv_counts = {}
    md_counts = {}

    # 1. 19 CSV из segments/
    segment_csvs = sorted((BASE_INTERVIEW / "segments").rglob("jobs.csv"))
    for p in segment_csvs:
        recs = parse_segment_csv(p)
        all_records.extend(recs)
        csv_counts[p.relative_to(BASE_INTERVIEW).as_posix()] = len(recs)
        print(f"[CSV-segment] {p.relative_to(BASE_INTERVIEW)}: {len(recs)} джобов")

    # 2. Агрегированный jobs_and_pains
    agg_path = BASE_INTERVIEW / "jobs_sources" / "jobs_and_pains" / "jobs.csv"
    agg_recs = parse_aggregated_csv(agg_path)
    all_records.extend(agg_recs)
    csv_counts[agg_path.relative_to(BASE_INTERVIEW).as_posix()] = len(agg_recs)
    print(f"[CSV-aggregated] jobs_and_pains/jobs.csv: {len(agg_recs)} джобов")

    # 3. MD JTBD — Catalina
    cat_path = BASE_KNOWLEDGE / "analysis_catalina_jtbd.md"
    if cat_path.exists():
        recs = parse_catalina(cat_path)
        all_records.extend(recs)
        md_counts[str(cat_path.relative_to(BASE_ROOT))] = len(recs)
        print(f"[MD-jtbd] {cat_path.name}: {len(recs)} джобов")

    # 4. MD JTBD — Padel
    pad_path = BASE_KNOWLEDGE / "Interviews/Padel/denis_padel_jtbd.md"
    if pad_path.exists():
        recs = parse_padel(pad_path)
        all_records.extend(recs)
        md_counts[str(pad_path.relative_to(BASE_ROOT))] = len(recs)
        print(f"[MD-jtbd] {pad_path.name}: {len(recs)} джобов")

    # 5. MD JTBD — sales_calls_1c_switch (5 файлов)
    sc_dir = BASE_KNOWLEDGE / "Interviews/sales_calls_1c_switch"
    sc_files = [
        ("01_anastasia_perm_franchise_exit.md", "S1", "J1", "Анастасия (Пермь) — выход из франшизы (1С+Mobi)"),
        ("02_evgeny_club_korp.md", "S2", "J2", "Евгений — клуб корпоративный (1С КОРП)"),
        ("03_kids_studio_judo_taekwondo.md", "S1", "J5", "Детская студия — дзюдо/тхэквондо (1С)"),
        ("04_anastasia_zavyalovo_micro_studio.md", "S5", "J1", "Анастасия (Завьялово) — микро-студия без админов"),
        ("05_marina_self_written_1c.md", "S2", "J2", "Марина — самописная 1С"),
    ]
    for fname, S, J, lbl in sc_files:
        p = sc_dir / fname
        if not p.exists():
            print(f"[MD-skip] нет файла: {p}")
            continue
        recs = parse_sales_call_1c(p, S, J, lbl)
        all_records.extend(recs)
        md_counts[str(p.relative_to(BASE_ROOT))] = len(recs)
        print(f"[MD-jtbd] {p.name}: {len(recs)} джобов")

    # 6. Narrative MD (только сводные/аналитические — оставлены как ссылка-плейсхолдер).
    # UrbanFit и Kochkin теперь имеют полные структурированные CSV выше, плейсхолдеры убраны.
    narrative_sources = [
        (BASE_KNOWLEDGE / "Interviews/Network/UrbanFit_Network_Analysis.md", "S3", "J4",
         "Urban Fit — сетевой анализ",
         "Аналитическая записка по сегменту, не структурированный JTBD. См. также полный CSV urbanfit_full."),
        (BASE_KNOWLEDGE / "Interviews/sales_calls_1c_switch/JTBD_summary_1c_segment.md", "?", "?",
         "Сводный JTBD по сегменту 1С (5 интервью + Каталина)",
         "Сводный документ по сегменту. Джобы выведены из 5 файлов sales_calls_1c_switch/, которые уже разобраны."),
    ]
    for p, S, J, lbl, note in narrative_sources:
        if not p.exists():
            continue
        recs = add_unparseable_reference(p, S, J, lbl, note)
        all_records.extend(recs)
        md_counts[str(p.relative_to(BASE_ROOT))] = len(recs)
        print(f"[MD-ref] {p.name}: {len(recs)} ссылка-плейсхолдер")

    # Нормализация размеров
    for r in all_records:
        r["job_size"] = normalize_size(r["job_size"])

    # Обнаружение дубликатов: группировка по нормализованному want_short
    from collections import defaultdict
    dup_groups = defaultdict(list)
    for r in all_records:
        if r["source_type"] == "md_jtbd_narrative":
            continue
        key = re.sub(r"[^а-яa-zё0-9 ]+", "", r["want_short"].lower()).strip()
        if not key:
            continue
        dup_groups[key].append(r)

    group_counter = 0
    for key, recs in dup_groups.items():
        if len(recs) < 2:
            continue
        group_counter += 1
        # Выбираем как canonical запись с самым полным контекстом (макс длина want_result + so_that + context_when)
        recs_sorted = sorted(recs, key=lambda x: -(len(x["want_result"]) + len(x["so_that"]) + len(x["context_when"]) + len(x["lpr_quote"])))
        canonical = recs_sorted[0]
        for r in recs:
            r["duplicate_group"] = f"D{group_counter:03d}"
            r["is_canonical"] = "1" if r is canonical else "0"
    # Не-дубликаты
    for r in all_records:
        if not r.get("duplicate_group"):
            r["duplicate_group"] = ""
            r["is_canonical"] = "1"

    # ID
    for i, r in enumerate(all_records, 1):
        r["id"] = f"J{i:03d}"

    # Только канонические записи — дубликаты в публичный вывод не попадают
    canonical_records = [r for r in all_records if r.get("is_canonical") == "1"]
    # Перенумеровываем ID после удаления дубликатов
    for i, r in enumerate(canonical_records, 1):
        r["id"] = f"J{i:03d}"

    out_csv = OUT / "jobs_master.csv"
    with open(out_csv, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=TARGET_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(canonical_records)

    # Verification report
    report = OUT / "verification_report.md"
    with open(report, "w", encoding="utf-8") as f:
        f.write("# Отчёт верификации сборки таблицы джобов\n\n")
        f.write(f"**Итого джобов в финальной таблице:** {len(all_records)}\n\n")
        f.write("## Источники\n\n")
        f.write("### CSV (Interview_Analysis)\n\n")
        f.write("| Файл | Извлечено |\n|------|-----------|\n")
        for k, v in csv_counts.items():
            f.write(f"| `{k}` | {v} |\n")
        f.write("\n### MD JTBD (Knowledge/Interviews)\n\n")
        f.write("| Файл | Извлечено |\n|------|-----------|\n")
        for k, v in md_counts.items():
            f.write(f"| `{k}` | {v} |\n")

        f.write("\n## Сводка по сегментам\n\n")
        seg_counts = defaultdict(int)
        for r in all_records:
            seg_counts[(r["business_segment"] or "?", r["job_segment"] or "?")] += 1
        f.write("| Бизнес-сегмент | Джоб-сегмент | Кол-во |\n|---|---|---|\n")
        for (s, j), c in sorted(seg_counts.items()):
            f.write(f"| {s} | {j} | {c} |\n")

        f.write("\n## Сводка по размеру\n\n")
        size_counts = defaultdict(int)
        for r in all_records:
            size_counts[r["job_size"] or "(не определён)"] += 1
        f.write("| Размер | Кол-во |\n|---|---|\n")
        for s, c in sorted(size_counts.items()):
            f.write(f"| {s} | {c} |\n")

        f.write("\n## Записи требующие внимания\n\n")
        problematic = [r for r in all_records if not r["want_short"]]
        if problematic:
            f.write(f"**Пустой want_short:** {len(problematic)} записей\n")
            for r in problematic[:30]:
                f.write(f"- `{r['id']}` из `{r['source_file']}` — {r.get('notes','(нет note)')}\n")
        else:
            f.write("Все записи имеют want_short.\n")

        # Дубликаты
        f.write("\n## ⚠️ Дубликаты\n\n")
        f.write(f"**Всего записей:** {len(all_records)}\n")
        f.write(f"**Канонических (после дедупликации):** {len(canonical_records)}\n")
        f.write(f"**Дубликатов:** {len(all_records) - len(canonical_records)}\n\n")
        dup_clusters = defaultdict(list)
        for r in all_records:
            if r.get("duplicate_group"):
                dup_clusters[r["duplicate_group"]].append(r)
        f.write(f"**Кластеров дубликатов:** {len(dup_clusters)}\n\n")
        f.write("### Топ-15 крупнейших кластеров дубликатов\n\n")
        clusters_sorted = sorted(dup_clusters.items(), key=lambda x: -len(x[1]))
        for grp, recs in clusters_sorted[:15]:
            sources = sorted(set(r["source_file"] for r in recs))
            f.write(f"- **{grp}** ({len(recs)} строк): `{recs[0]['want_short'][:80]}`\n")
            for s in sources:
                f.write(f"  - `{s}`\n")
        if len(clusters_sorted) > 15:
            f.write(f"\n*(ещё {len(clusters_sorted)-15} кластеров меньшего размера)*\n")

        f.write("\n### Рекомендация\n\n")
        f.write("Файлы dmitry_network_* (3 шт) и большинство файлов switch_other_systems/ (10 шт)\n")
        f.write("содержат идентичный контент — это итерации обработки одного интервью\n")
        f.write("(Дмитрий и Анна/танцы соответственно). Используйте `jobs_master_deduplicated.csv`\n")
        f.write("для уникальных джобов или фильтруйте `is_canonical=1` в основном файле.\n")

    print(f"\n✅ CSV: {out_csv}")
    print(f"✅ Отчёт: {report}")
    print(f"📊 Итого джобов: {len(canonical_records)} (дубликаты исключены)")
    return canonical_records


if __name__ == "__main__":
    main()
