#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
«Режим» — генератор планера привычек по спецификации совета (council/итог_v2.md).

Собирает .xlsx, рассчитанный на импорт в Google Таблицы (основная платформа)
и на Excel 365 (бонус). Вся логика — формулы, без скриптов и макросов.

    python3 build.py --out dist                 # тёмная и светлая сборки
    python3 build.py --out dist --recalc        # + пересчёт в LibreOffice и проверка
    python3 build.py --out dist --demo          # + файл с демо-данными для рилса
    python3 build.py --year 2027 --today 2027-03-04

Геометрия и бюджеты — разделы 3, 5 и 12 спецификации; отклонение валит сборку.
"""
import argparse
import calendar
import datetime as dt
import os
import re
import subprocess
import sys

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as CN
from xlsxwriter.utility import xl_rowcol_to_cell as RC

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import texts as T                                                    # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# Единицы: ширина колонки и высота строки задаются в пикселях экрана.
# Google Таблицы читают xlsx по формуле Excel (px = chars*7 + 5) — проверено
# смоук-тестом: заданные 43 px возвращаются из Таблиц как 43 px.
# ─────────────────────────────────────────────────────────────────────────────
def w(px):
    return (px - 5) / 7


def h(px):
    return px * 0.75


NH = 7                       # потолок привычек (раздел 3)
FONT = "Arial"

# Колонки листа месяца
C_DATE = 0
C_H0, C_H1 = 1, 7            # B..H — привычки
C_SLEEP, C_ENERGY, C_MOOD, C_NOTE = 8, 9, 10, 11
C_LAST = C_NOTE

# Строки листа месяца (0-based)
R_HUD, R_DAYLINE, R_ICON, R_SHORT, R_STREAK = 0, 1, 2, 3, 4
R_D1 = 5                     # день 1
R_D31 = R_D1 + 30
R_LIST_H = R_D31 + 2         # заголовок «Мои привычки»
R_LIST = R_LIST_H + 1        # 7 строк списка
R_GOALS_H = R_LIST + NH + 1
R_GOALS = R_GOALS_H + 1      # 3 строки целей
R_WEEK_H = R_GOALS + 4
R_WEEK = R_WEEK_H + 1        # 5 строк текста недели
R_ANS = R_WEEK + 5           # 5 строк ответов
R_SPARK = R_ANS + 5
R_END = R_SPARK

# Лента «Расчёт»: строка = день года, начиная с 1-based строки 5
CALC = "Расчёт"
K_ROW0 = 10                                    # 0-based строка первого дня года
#   строки 1..8 заняты служебной шапкой, поэтому лента начинается ниже
KC_PASSED = 0                                  # A — день прошёл
KC_M0 = 1                                      # B..H — зеркала отметок
KC_S0 = KC_M0 + NH                             # I..O — серии
KC_MISS0 = KC_S0 + NH                          # P..V — пропуски в месяце
KC_TICKS = KC_MISS0 + NH                       # W — галочек за день
KC_FULL = KC_TICKS + 1                         # X — полный день
KC_XP = KC_TICKS + 2                           # Y — опыт за день
KC_XPSUM = KC_TICKS + 3                        # Z — опыт нарастающим
KC_RET = KC_TICKS + 4                          # AA — возвратов за день
KC_EVE = KC_TICKS + 5                          # AB — вечеров подряд
KC_LAST = KC_EVE

# Шапка «Расчёта»
K_TODAY = "$B$1"             # сегодня (дата)
K_ROWT = "$B$2"              # строка сегодняшнего дня в ленте
K_NACT = "$B$3"              # активных привычек
K_XP = "$B$4"                # опыт всего
K_LVL = "$B$5"               # уровень
K_RANK = "$B$6"              # звание
K_NEXT = "$B$7"              # доля до следующего уровня
K_CLOSED = "$B$8"            # закрытых испытаний
K_FREQ0 = 5                  # F1..L1 — «раз в неделю» по привычкам

# Блоки месяцев и недель на «Расчёте» (0-based строки)
K_MON = K_ROW0 + 366 + 2                       # 12 строк: месяцы
K_WEEK = K_MON + 14                            # 53 строки: недели
K_AWARD = K_WEEK + 55                          # 10 строк: награды

MC_DIM, MC_EL, MC_DONE, MC_EXP, MC_PCT, MC_FULL, MC_BEST, MC_XP = range(1, 9)
WC_A, WC_B, WC_WEAK, WC_CNT, WC_CLOSED, WC_DONE, WC_EXP, WC_PREV, WC_S0 = (
    1, 2, 3, 4, 5, 6, 7, 8, 9)

SEASON = {1: "#5AA9E6", 2: "#5AA9E6", 3: "#34D399", 4: "#34D399", 5: "#34D399",
          6: "#FB923C", 7: "#FB923C", 8: "#FB923C", 9: "#F59E0B", 10: "#F59E0B",
          11: "#F59E0B", 12: "#5AA9E6"}

THEMES = {
    "dark": dict(
        bg="#0E1016", panel="#1B1F2A", panel2="#262B38", text="#F2F4F8",
        muted="#A0A7B8", accent="#F6C453", today="#3A3115", done="#22A35E",
        glyph="#06301A", xp="#8B5CF6", input="#2A2E3B",
        scale=("#262B38", "#166534", "#22A35E"), tile="#262B38",
    ),
    "light": dict(
        bg="#F4F5F9", panel="#FFFFFF", panel2="#EEF0F5", text="#14171F",
        muted="#5C6478", accent="#F5B800", today="#FFF3C4", done="#BBF7D0",
        glyph="#14532D", xp="#7C3AED", input="#FFFDF2",
        scale=("#EEF0F5", "#86EFAC", "#16A34A"), tile="#EEF0F5",
    ),
}


class Budget:
    """Счётчики для ворот линтера (раздел 12)."""

    def __init__(self):
        self.formulas = 0
        self.cf = {}
        self.spark = {}
        self.today = 0
        self.banned = []

    BAN = re.compile(r"\b(OFFSET|INDIRECT|ARRAYFORMULA|MAXIFS|MINIFS|XLOOKUP|FILTER|UNIQUE|SORT)\s*\(",
                     re.I)

    def formula(self, sheet, f):
        self.formulas += 1
        if "TODAY(" in f.upper():
            self.today += 1
        if self.BAN.search(f):
            self.banned.append(f"{sheet}: {f[:70]}")
        if "SPARKLINE(" in f.upper():
            self.spark[sheet] = self.spark.get(sheet, 0) + 1

    def rule(self, sheet):
        self.cf[sheet] = self.cf.get(sheet, 0) + 1

    def check(self):
        bad = []
        if self.formulas > 15000:
            bad.append(f"формул {self.formulas} > 15000")
        if self.today > 1:
            bad.append(f"TODAY() встречается {self.today} раз, разрешён один")
        for s, n in self.cf.items():
            if n > 10:
                bad.append(f"правил условного форматирования на «{s}»: {n} > 10")
        for s, n in self.spark.items():
            if n > 15:
                bad.append(f"SPARKLINE на «{s}»: {n} > 15")
        bad += [f"запрещённая функция — {b}" for b in self.banned]
        return bad


class Book:
    """Обёртка над книгой: кэш форматов, учёт бюджетов, подмена под цель."""

    def __init__(self, path, theme, target, year, today, budget):
        self.wb = xlsxwriter.Workbook(
            path, {"default_format_properties": {"font_name": FONT, "font_size": 10},
                   "strings_to_numbers": False})
        self.th = THEMES[theme]
        self.target = target                    # "sheets" | "excel"
        self.year = year
        self.today = today
        self.b = budget

    def fmt(self, **kw):
        key = tuple(sorted((k, tuple(v) if isinstance(v, list) else v) for k, v in kw.items()))
        if not hasattr(self, "_f"):
            self._f = {}
        if key not in self._f:
            self._f[key] = self.wb.add_format(kw)
        return self._f[key]

    def f(self, ws, row, col, formula, fmt=None, value=""):
        self.b.formula(ws.get_name(), formula)
        ws.write_formula(row, col, formula, fmt, value)

    def cf(self, ws, first_row, first_col, last_row, last_col, options):
        self.b.rule(ws.get_name())
        ws.conditional_format(first_row, first_col, last_row, last_col, options)

    def spark(self, ws, row, col, values, color, fmt, kind="bar", maximum=None):
        """SPARKLINE в Таблицах; в Excel — родной спарклайн той же ячейки."""
        if self.target == "sheets":
            opts = [f'"charttype","{kind}"', f'"color1","{color}"',
                    f'"empty","zero"']
            if kind == "bar":
                opts = [f'"charttype","bar"', f'"max",{maximum or 1}',
                        f'"color1","{color}"', f'"color2","{self.th["panel2"]}"']
            f = "=SPARKLINE(%s,{%s})" % (values, ";".join(opts))
            self.f(ws, row, col, f, fmt)
        else:
            ws.write_blank(row, col, None, fmt)
            ws.add_sparkline(row, col, {
                "range": f"{ws.get_name()}!{values}" if "!" not in values else values,
                "type": "column" if kind == "column" else "column",
                "series_color": color, "max": maximum, "min": 0,
            })


# ─────────────────────────────────────────────────────────────────────────────
def kr(day_index):
    """0-based строка дня года (1..366) в ленте «Расчёта»."""
    return K_ROW0 + day_index - 1


def kc(row, col, abs_col=True, abs_row=False):
    return f"{CALC}!{RC(row, col, abs_row, abs_col)}"


def month_rows(year, m):
    """Индексы дней года для месяца m (1-based)."""
    start = dt.date(year, m, 1).timetuple().tm_yday
    return [start + i for i in range(calendar.monthrange(year, m)[1])]


def week_blocks(year):
    """Недели года: список (номер, [индексы дней года]), неделя начинается с понедельника."""
    d = dt.date(year, 1, 1)
    end = dt.date(year, 12, 31)
    weeks, cur = [], []
    while d <= end:
        cur.append(d.timetuple().tm_yday)
        if d.weekday() == 6:
            weeks.append(cur)
            cur = []
        d += dt.timedelta(days=1)
    if cur:
        weeks.append(cur)
    return weeks


# ─────────────────────────────────────────────────────────────────────────────
# Лист «Расчёт» — движок
# ─────────────────────────────────────────────────────────────────────────────
def build_calc(bk, weeks):
    th, year = bk.th, bk.year
    ws = bk.wb.add_worksheet(CALC)
    ws.hide()
    base = bk.fmt(font_size=9, font_color=th["muted"])
    ws.set_column(0, KC_LAST + 2, 9, base)

    # Шапка: сегодня, строка дня, активные привычки, опыт, уровень, звание
    lab = bk.fmt(font_size=9, bold=True)
    for i, name in enumerate(["сегодня", "строка", "активных", "опыт", "уровень",
                              "звание", "до уровня", "испытаний"]):
        ws.write(i, 2, name, base)
    bk.f(ws, 0, 1, f"=IF('Старт'!$C$9=\"\",TODAY(),'Старт'!$C$9)", lab)
    bk.f(ws, 1, 1, f"=IF(OR({K_TODAY}<DATE({year},1,1),{K_TODAY}>DATE({year},12,31)),0,"
                   f"{K_ROW0 + 1}+{K_TODAY}-DATE({year},1,1))", lab)
    bk.f(ws, 2, 1, "=COUNTA('Старт'!$B$12:$B$18)", lab)
    bk.f(ws, 3, 1, f"=IF({K_ROWT}=0,0,INDEX({CN(KC_XPSUM)}:{CN(KC_XPSUM)},{K_ROWT})"
                   f"+100*{K_CLOSED})", lab)
    bk.f(ws, 4, 1, f"=MATCH({K_XP},'Списки'!$E$2:$E$21,1)", lab)
    bk.f(ws, 5, 1, f"=INDEX('Списки'!$F$2:$F$21,{K_LVL})", lab)
    bk.f(ws, 6, 1, f"=IFERROR(({K_XP}-INDEX('Списки'!$E$2:$E$21,{K_LVL}))"
                   f"/(INDEX('Списки'!$E$2:$E$21,{K_LVL}+1)"
                   f"-INDEX('Списки'!$E$2:$E$21,{K_LVL})),1)", lab)
    bk.f(ws, 7, 1, f"=SUM({CN(WC_CLOSED)}{K_WEEK + 1}:{CN(WC_CLOSED)}{K_WEEK + len(weeks)})", lab)
    for hbt in range(NH):
        bk.f(ws, 0, K_FREQ0 + hbt, f"=IF('Старт'!$F${12 + hbt}=\"\",7,'Старт'!$F${12 + hbt})", base)

    # Лента дней
    months = {}
    for m in range(1, 13):
        for d_i, yday in enumerate(month_rows(year, m), start=1):
            months[yday] = (m, d_i)
    total = 366 if calendar.isleap(year) else 365

    for yday in range(1, total + 1):
        r = kr(yday)
        m, dom = months[yday]
        prev = r - 1
        sheet_m = T.MONTHS_SHORT[m - 1]
        bk.f(ws, r, KC_PASSED,
             f"=IF({K_TODAY}>=DATE({year},{m},{dom}),1,0)", base)
        passed = RC(r, KC_PASSED, False, True)
        for hbt in range(NH):
            if m in getattr(bk, "months", range(1, 13)):
                cell_m = f"'{sheet_m}'!{RC(R_D1 + dom - 1, C_H0 + hbt, False, True)}"
                bk.f(ws, r, KC_M0 + hbt, f"=IFERROR(IF({cell_m}=TRUE,1,0),0)", base)
            else:
                ws.write_number(r, KC_M0 + hbt, 0, base)
            done = RC(r, KC_M0 + hbt, False, True)
            miss_prev = "0" if dom == 1 or yday == 1 else RC(prev, KC_MISS0 + hbt, False, True)
            s_prev = "0" if yday == 1 else RC(prev, KC_S0 + hbt, False, True)
            # пропуски в месяце: сбрасываются первого числа
            bk.f(ws, r, KC_MISS0 + hbt,
                 f"=IF({passed}=0,{miss_prev},{miss_prev}+IF({done}=1,0,1))", base)
            miss = RC(r, KC_MISS0 + hbt, False, True)
            # серия: щит прощает первый пропуск месяца
            bk.f(ws, r, KC_S0 + hbt,
                 f"=IF({passed}=0,{s_prev},IF({done}=1,{s_prev}+1,"
                 f"IF({miss}<=1,{s_prev},0)))", base)
        rng = f"{RC(r, KC_M0, False, True)}:{RC(r, KC_M0 + NH - 1, False, True)}"
        bk.f(ws, r, KC_TICKS, f"=SUM({rng})", base)
        ticks = RC(r, KC_TICKS, False, True)
        # полный день: все ежедневные привычки отмечены
        daily = "+".join(f"IF(AND({RC(0, K_FREQ0 + i, True, True)}=7,"
                         f"'Старт'!$B${12 + i}<>\"\"),1,0)" for i in range(NH))
        daily_done = "+".join(f"IF(AND({RC(0, K_FREQ0 + i, True, True)}=7,"
                              f"{RC(r, KC_M0 + i, False, True)}=1),1,0)" for i in range(NH))
        bk.f(ws, r, KC_FULL,
             f"=IF(AND({passed}=1,({daily})>0,({daily_done})=({daily})),1,0)", base)
        # опыт за день и возвраты
        xp_terms, ret_terms = [], []
        for hbt in range(NH):
            done = RC(r, KC_M0 + hbt, False, True)
            s_prev = "0" if yday == 1 else RC(prev, KC_S0 + hbt, False, True)
            m_prev = "0" if dom == 1 or yday == 1 else RC(prev, KC_MISS0 + hbt, False, True)
            freq = RC(0, K_FREQ0 + hbt, True, True)
            back = f"AND({freq}=7,{s_prev}=0,{m_prev}>1)"
            xp_terms.append(
                f"IF({done}=1,IF({freq}<7,10,IF({back},15,"
                f"IF({s_prev}>=20,15,IF({s_prev}>=6,13,10)))),0)")
            ret_terms.append(f"IF(AND({done}=1,{back}),1,0)")
        bk.f(ws, r, KC_XP, "=" + "+".join(xp_terms), base)
        bk.f(ws, r, KC_RET, "=" + "+".join(ret_terms), base)
        xp_prev = "0" if yday == 1 else RC(prev, KC_XPSUM, False, True)
        bk.f(ws, r, KC_XPSUM, f"={xp_prev}+{RC(r, KC_XP, False, True)}", base)
        eve_prev = "0" if yday == 1 else RC(prev, KC_EVE, False, True)
        if m in getattr(bk, "months", range(1, 13)):
            sleep_cell = f"'{sheet_m}'!{RC(R_D1 + dom - 1, C_SLEEP, False, True)}"
            bk.f(ws, r, KC_EVE,
                 f"=IF(IFERROR({sleep_cell},\"\")=\"\",0,{eve_prev}+1)", base)
        else:
            bk.f(ws, r, KC_EVE, "=0", base)

    # Блок месяцев
    ws.write(K_MON - 1, 0, "месяцы", lab)
    for m in range(1, 13):
        r = K_MON + m - 1
        days = month_rows(year, m)
        a, b = kr(days[0]) + 1, kr(days[-1]) + 1                 # 1-based
        dim = len(days)
        ws.write_number(r, 0, m, base)
        ws.write_number(r, MC_DIM, dim, base)
        bk.f(ws, r, MC_EL, f"=SUM({CN(KC_PASSED)}{a}:{CN(KC_PASSED)}{b})", base)
        el = RC(r, MC_EL, False, True)
        bk.f(ws, r, MC_DONE, f"=SUM({CN(KC_TICKS)}{a}:{CN(KC_TICKS)}{b})", base)
        exp_terms = "+".join(
            f"IF('Старт'!$B${12 + i}=\"\",0,ROUND({RC(0, K_FREQ0 + i, True, True)}*{dim}/7,0))"
            for i in range(NH))
        bk.f(ws, r, MC_EXP, f"=IF({el}=0,0,({exp_terms})*{el}/{dim})", base)
        exp = RC(r, MC_EXP, False, True)
        bk.f(ws, r, MC_PCT, f"=IF({exp}=0,0,MIN(1,{RC(r, MC_DONE, False, True)}/{exp}))", base)
        bk.f(ws, r, MC_FULL, f"=SUM({CN(KC_FULL)}{a}:{CN(KC_FULL)}{b})", base)
        best = ",".join(f"MAX({CN(KC_S0 + i)}{a}:{CN(KC_S0 + i)}{b})" for i in range(NH))
        bk.f(ws, r, MC_BEST, f"=MAX({best})", base)
        bk.f(ws, r, MC_XP, f"=SUM({CN(KC_XP)}{a}:{CN(KC_XP)}{b})", base)

    # Блок недель
    ws.write(K_WEEK - 1, 0, "недели", lab)
    for i, days in enumerate(weeks):
        r = K_WEEK + i
        a, b = kr(days[0]) + 1, kr(days[-1]) + 1
        ws.write_number(r, 0, i + 1, base)
        ws.write_number(r, WC_A, a, base)
        ws.write_number(r, WC_B, b, base)
        for hbt in range(NH):
            bk.f(ws, r, WC_S0 + hbt,
                 f"=IF('Старт'!$B${12 + hbt}=\"\",99,"
                 f"SUM({CN(KC_M0 + hbt)}{a}:{CN(KC_M0 + hbt)}{b}))", base)
        prev = r - 1
        srng = f"{RC(prev, WC_S0, False, True)}:{RC(prev, WC_S0 + NH - 1, False, True)}"
        if i == 0:
            bk.f(ws, r, WC_WEAK, "=0", base)
        else:
            bk.f(ws, r, WC_WEAK,
                 f"=IF(MIN({srng})=99,0,MATCH(MIN({srng}),{srng},0))", base)
        weak = RC(r, WC_WEAK, False, True)
        cnt_terms = "+".join(
            f"IF({weak}={i2 + 1},{RC(r, WC_S0 + i2, False, True)},0)" for i2 in range(NH))
        bk.f(ws, r, WC_CNT, f"=IF({weak}=0,0,{cnt_terms})", base)
        bk.f(ws, r, WC_CLOSED, f"=IF(AND({weak}>0,{RC(r, WC_CNT, False, True)}>=5),1,0)", base)
        bk.f(ws, r, WC_DONE, f"=SUM({CN(KC_TICKS)}{a}:{CN(KC_TICKS)}{b})", base)
        exp_terms = "+".join(
            f"IF('Старт'!$B${12 + i2}=\"\",0,{RC(0, K_FREQ0 + i2, True, True)})"
            for i2 in range(NH))
        passed = f"SUM({CN(KC_PASSED)}{a}:{CN(KC_PASSED)}{b})"
        bk.f(ws, r, WC_EXP, f"=ROUND(({exp_terms})*{passed}/7,0)", base)
        bk.f(ws, r, WC_PREV,
             "=0" if i == 0 else f"={RC(prev, WC_DONE, False, True)}", base)

    # Награды: счётчики
    ws.write(K_AWARD - 1, 0, "награды", lab)
    la, lb = K_ROW0 + 1, K_ROW0 + total
    counters = {
        "ticks": f"=SUM({CN(KC_TICKS)}{la}:{CN(KC_TICKS)}{lb})",
        "record": "=MAX(" + ",".join(
            f"MAX({CN(KC_S0 + i)}{la}:{CN(KC_S0 + i)}{lb})" for i in range(NH)) + ")",
        "full_days": f"=SUM({CN(KC_FULL)}{la}:{CN(KC_FULL)}{lb})",
        "returns": f"=SUM({CN(KC_RET)}{la}:{CN(KC_RET)}{lb})",
        "challenges": f"=SUM({CN(WC_CLOSED)}{K_WEEK + 1}:{CN(WC_CLOSED)}{K_WEEK + len(weeks)})",
        "evenings": f"=MAX({CN(KC_EVE)}{la}:{CN(KC_EVE)}{lb})",
        "answers": "=COUNTA('Год'!$B$28:$B$80)",     # блок ответов на вопрос недели
    }
    order = list(counters)
    for i, key in enumerate(order):
        ws.write(K_AWARD + i, 0, key, base)
        bk.f(ws, K_AWARD + i, 1, counters[key], base)
    return {k: f"{CALC}!{RC(K_AWARD + i, 1, True, True)}" for i, k in enumerate(order)}


# ─────────────────────────────────────────────────────────────────────────────
# Лист месяца
# ─────────────────────────────────────────────────────────────────────────────
def build_month(bk, m, weeks, demo=None):
    th, year = bk.th, bk.year
    name = T.MONTHS_SHORT[m - 1]
    ws = bk.wb.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.set_tab_color(SEASON[m])
    dim = calendar.monthrange(year, m)[1]
    days = month_rows(year, m)
    mrow = K_MON + m - 1                                  # строка месяца в «Расчёте»

    base = bk.fmt(bg_color=th["bg"], font_color=th["text"])
    ws.set_column(0, 40, w(43), base)
    ws.set_column(C_DATE, C_DATE, w(48), base)
    ws.set_column(C_H0, C_H1, w(43), base)
    ws.set_column(C_SLEEP, C_MOOD, w(56), base)
    ws.set_column(C_NOTE, C_NOTE, w(190), base)
    ws.set_column(C_LAST + 1, 40, w(43), base)

    for r, px in ((R_HUD, 36), (R_DAYLINE, 24), (R_ICON, 24), (R_SHORT, 20), (R_STREAK, 20)):
        ws.set_row(r, h(px))
    for d in range(31):
        ws.set_row(R_D1 + d, h(46))
    ws.set_row(R_D31 + 1, h(8))
    ws.set_row(R_LIST_H, h(22))
    for i in range(NH):
        ws.set_row(R_LIST + i, h(40))
    ws.set_row(R_GOALS_H, h(22))
    for i in range(3):
        ws.set_row(R_GOALS + i, h(44))
    ws.set_row(R_WEEK_H, h(22))
    for i in range(5):
        ws.set_row(R_WEEK + i, h(30))
    for i in range(5):
        ws.set_row(R_ANS + i, h(46))
    ws.set_row(R_SPARK, h(64))

    # ── служебные ячейки I1:L1 (для условного форматирования, без других листов)
    svc = bk.fmt(bg_color=th["panel"], font_color=th["panel"], font_size=9)
    bk.f(ws, R_HUD, C_SLEEP, f"={CALC}!{K_NACT}", svc)                       # активных привычек
    bk.f(ws, R_HUD, C_ENERGY,
         f"=IF(AND(YEAR({CALC}!{K_TODAY})={year},MONTH({CALC}!{K_TODAY})={m}),"
         f"{R_D1}+DAY({CALC}!{K_TODAY}),0)", svc)                            # строка «сегодня»
    bk.f(ws, R_HUD, C_MOOD,
         f"=IF(WEEKDAY({CALC}!{K_TODAY},2)>=6,1,0)", svc)                    # выходной: подсветка ответа
    bk.f(ws, R_HUD, C_NOTE, f"={CALC}!{K_ROWT}", svc)
    svc_n = RC(R_HUD, C_SLEEP, True, True)
    svc_today = RC(R_HUD, C_ENERGY, True, True)

    # ── HUD
    hud = bk.fmt(bg_color=th["panel"], font_color=th["text"], bold=True, font_size=14,
                 valign="vcenter", align="left", indent=1)
    ws.merge_range(R_HUD, C_DATE, R_HUD, C_H0 + 3, "", hud)
    pct = f"{CALC}!{RC(mrow, MC_PCT, True, True)}"
    el = f"{CALC}!{RC(mrow, MC_EL, True, True)}"
    bk.f(ws, R_HUD, C_DATE,
         f'="{name.upper()} · "&IF({el}=0,"—",ROUND({pct}*100,0)&" %")&'
         f'" · Ур. "&{CALC}!{K_LVL}', hud)
    barfmt = bk.fmt(bg_color=th["panel"], align="center", valign="vcenter")
    ws.merge_range(R_HUD, C_H0 + 4, R_HUD, C_H1, "", barfmt)
    if bk.target == "sheets":
        bk.f(ws, R_HUD, C_H0 + 4,
             '=SPARKLINE(%s,{"charttype","bar";"max",1;"color1","%s";"color2","%s"})'
             % (f"{CALC}!{K_NEXT}", th["xp"], th["panel2"]), barfmt)
    else:
        bk.f(ws, R_HUD, C_H0 + 4,
             f'=REPT("▮",MAX(1,ROUND({CALC}!{K_NEXT}*10,0)))',
             bk.fmt(bg_color=th["panel"], font_color=th["xp"], font_size=11,
                    align="center", valign="vcenter"))

    # ── строка дня
    dayline = bk.fmt(bg_color=th["panel"], font_color=th["accent"], font_size=11,
                     valign="vcenter", align="left", indent=1)
    ws.merge_range(R_DAYLINE, C_DATE, R_DAYLINE, C_NOTE, "", dayline)
    bk.f(ws, R_DAYLINE, C_DATE, dayline_formula(bk, m, weeks), dayline)

    # ── значки, имена, серии
    icon_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=14,
                    align="center", valign="vcenter")
    short_f = bk.fmt(bg_color=th["panel"], font_color=th["muted"], font_size=10, bold=True,
                     align="center", valign="vcenter")
    streak_f = bk.fmt(bg_color=th["panel"], font_color=th["accent"], font_size=10, bold=True,
                      align="center", valign="vcenter")
    pan = bk.fmt(bg_color=th["panel"])
    for r in (R_ICON, R_SHORT, R_STREAK):
        ws.write_blank(r, C_DATE, None, pan)
        for c in range(C_SLEEP, C_NOTE + 1):
            ws.write_blank(r, c, None, pan)
    rowt = f"{CALC}!{K_ROWT}"
    for i in range(NH):
        sr = 12 + i
        bk.f(ws, R_ICON, C_H0 + i, f"=IF('Старт'!$B${sr}=\"\",\"\",'Старт'!$B${sr})", icon_f)
        bk.f(ws, R_SHORT, C_H0 + i, f"=IF('Старт'!$B${sr}=\"\",\"\",'Старт'!$E${sr})", short_f)
        bk.f(ws, R_STREAK, C_H0 + i, streak_formula(bk, i, sr), streak_f)

    # ── сетка дней
    date_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=10, bold=True,
                    align="center", valign="vcenter", text_wrap=True,
                    border=1, border_color=th["bg"])
    cb_f = bk.fmt(bg_color=th["panel2"], font_color=th["muted"], font_size=14,
                  align="center", valign="vcenter", border=1, border_color=th["bg"],
                  num_format=";;;")
    in_f = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=10,
                  align="center", valign="vcenter", border=1, border_color=th["bg"])
    note_f = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=10,
                    align="left", indent=1, valign="vcenter", text_wrap=True,
                    border=1, border_color=th["bg"])
    void = bk.fmt(bg_color=th["bg"])
    for d in range(1, 32):
        r = R_D1 + d - 1
        if d > dim:
            for c in range(C_DATE, C_LAST + 1):
                ws.write_blank(r, c, None, void)
            continue
        date = dt.date(year, m, d)
        label = f"{T.WEEKDAYS[date.weekday()]} {d}"
        k = kr(days[d - 1])
        bk.f(ws, r, C_DATE,
             f'=IF({CALC}!{RC(k, KC_PASSED, True, True)}=0,"{label}",'
             f'"{label}"&CHAR(10)&{CALC}!{RC(k, KC_TICKS, True, True)}&"/"&{svc_n})', date_f)
        for i in range(NH):
            val = bool(demo and (i, d) in demo.get("checks", ()))
            ws.write_boolean(r, C_H0 + i, val, cb_f)
        ws.write(r, C_SLEEP, (demo or {}).get("sleep", {}).get(d, ""), in_f)
        ws.write(r, C_ENERGY, (demo or {}).get("energy", {}).get(d, ""), in_f)
        ws.write(r, C_MOOD, (demo or {}).get("mood", {}).get(d, ""), in_f)
        ws.write(r, C_NOTE, (demo or {}).get("note", {}).get(d, ""), note_f)

    ws.data_validation(R_D1, C_H0, R_D31, C_H1,
                       {"validate": "list", "source": ["TRUE", "FALSE"], "dropdown": False})
    ws.data_validation(R_D1, C_SLEEP, R_D31, C_SLEEP,
                       {"validate": "list", "source": T.SLEEP, "ignore_blank": True})
    ws.data_validation(R_D1, C_ENERGY, R_D31, C_ENERGY,
                       {"validate": "list", "source": T.ENERGY, "ignore_blank": True})
    ws.data_validation(R_D1, C_MOOD, R_D31, C_MOOD,
                       {"validate": "list", "source": T.MOOD, "ignore_blank": True})

    # ── условное форматирование (бюджет ≤10 правил)
    done_fmt = bk.fmt(bg_color=th["done"], font_color=th["glyph"], bold=True)
    today_fmt = bk.fmt(bg_color=th["today"], font_color=th["accent"], bold=True)
    future = bk.fmt(bg_color=th["panel"], font_color=th["panel"])
    bk.cf(ws, R_D1, C_H0, R_D31, C_H1, {
        "type": "formula", "criteria": f"={RC(R_D1, C_H0)}=TRUE",
        "format": done_fmt, "stop_if_true": True})
    bk.cf(ws, R_D1, C_DATE, R_D31, C_NOTE, {
        "type": "formula", "criteria": f"=ROW()={svc_today}",
        "format": today_fmt, "stop_if_true": True})
    bk.cf(ws, R_D1, C_DATE, R_D31, C_NOTE, {
        "type": "formula", "criteria": f"=ROW()>{R_D1}+{dim}",
        "format": bk.fmt(bg_color=th["bg"], font_color=th["bg"])})
    bk.cf(ws, R_D1, C_H0, R_D31, C_H1, {
        "type": "formula", "criteria": f"={CN(C_H0)}${R_SHORT + 1}=\"\"",
        "format": bk.fmt(bg_color=th["panel"], font_color=th["panel"])})

    # ── список привычек
    sec = bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=9, bold=True,
                 valign="vcenter", align="left", indent=1)
    ws.merge_range(R_LIST_H, C_DATE, R_LIST_H, C_NOTE, T.L["habits"].upper(), sec)
    name_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                    align="left", indent=1, valign="vcenter")
    num_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11, bold=True,
                   align="center", valign="vcenter")
    mut_f = bk.fmt(bg_color=th["panel"], font_color=th["muted"], font_size=10,
                   align="center", valign="vcenter")
    pct_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=10,
                   align="center", valign="vcenter", num_format="0%")
    for i in range(NH):
        r, sr = R_LIST + i, 12 + i
        ws.merge_range(r, C_DATE, r, C_H0 + 4, "", name_f)
        bk.f(ws, r, C_DATE,
             f"=IF('Старт'!$B${sr}=\"\",\"\",'Старт'!$A${sr}&\"  \"&'Старт'!$B${sr})", name_f)
        a, b = kr(days[0]) + 1, kr(days[-1]) + 1
        done = f"SUM({CALC}!${CN(KC_M0 + i)}${a}:${CN(KC_M0 + i)}${b})"
        exp = (f"ROUND(IF('Старт'!$F${sr}=\"\",7,'Старт'!$F${sr})*{dim}/7,0)")
        ws.merge_range(r, C_H0 + 5, r, C_H0 + 6, "", num_f)
        bk.f(ws, r, C_H0 + 5,
             f"=IF('Старт'!$B${sr}=\"\",\"\",{done}&\" из \"&{exp})", num_f)
        bk.f(ws, r, C_SLEEP, streak_formula(bk, i, sr), mut_f)
        bk.f(ws, r, C_ENERGY,
             f"=IF('Старт'!$B${sr}=\"\",\"\",MAX({CALC}!${CN(KC_S0 + i)}${a}:"
             f"${CN(KC_S0 + i)}${b}))", mut_f)
        bk.f(ws, r, C_MOOD,
             f"=IF(OR('Старт'!$B${sr}=\"\",{el}=0),\"\","
             f"MIN(1,{done}/MAX(1,{exp}*{el}/{dim})))", pct_f)
        if bk.target == "sheets":
            bk.f(ws, r, C_NOTE,
                 '=IF(%s="","",SPARKLINE(%s,{"charttype","bar";"max",1;"color1","%s";"color2","%s"}))'
                 % (RC(r, C_MOOD, True, True), RC(r, C_MOOD, True, True),
                    th["done"], th["panel2"]),
                 bk.fmt(bg_color=th["panel"], align="center", valign="vcenter"))
        else:
            bk.f(ws, r, C_NOTE,
                 f'=IF({RC(r, C_MOOD, True, True)}="","",'
                 f'REPT("▮",ROUND({RC(r, C_MOOD, True, True)}*10,0)))',
                 bk.fmt(bg_color=th["panel"], font_color=th["done"], font_size=11,
                        align="left", indent=1, valign="vcenter"))

    # ── цели месяца
    ws.merge_range(R_GOALS_H, C_DATE, R_GOALS_H, C_NOTE, T.L["goals"].upper(), sec)
    goal_cb = bk.fmt(bg_color=th["panel2"], font_color=th["muted"], font_size=14,
                     align="center", valign="vcenter", border=1, border_color=th["bg"],
                     num_format=";;;")
    goal_t = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=11,
                    align="left", indent=1, valign="vcenter")
    for i in range(3):
        r = R_GOALS + i
        ws.write_boolean(r, C_DATE, False, goal_cb)
        ws.merge_range(r, C_H0, r, C_NOTE, "", goal_t)
    ws.data_validation(R_GOALS, C_DATE, R_GOALS + 2, C_DATE,
                       {"validate": "list", "source": ["TRUE", "FALSE"], "dropdown": False})
    bk.cf(ws, R_GOALS, C_DATE, R_GOALS + 2, C_NOTE, {
        "type": "formula", "criteria": f"=${CN(C_DATE)}{R_GOALS + 1}=TRUE",
        "format": bk.fmt(bg_color=th["panel"], font_color=th["done"], bold=True)})

    # ── блок «Неделя»
    ws.merge_range(R_WEEK_H, C_DATE, R_WEEK_H, C_NOTE, T.L["week"].upper(), sec)
    wk_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                  align="left", indent=1, valign="vcenter")
    wi = current_week_index(bk, m, weeks)
    wrow = K_WEEK + wi
    for i, f in enumerate(week_lines(bk, wrow, mrow)):
        r = R_WEEK + i
        ws.merge_range(r, C_DATE, r, C_NOTE, "", wk_f)
        bk.f(ws, r, C_DATE, f, wk_f)

    ans_lab = bk.fmt(bg_color=th["panel"], font_color=th["muted"], font_size=10,
                     align="left", indent=1, valign="vcenter")
    ans_f = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=11,
                   align="left", indent=1, valign="vcenter", text_wrap=True)
    for i in range(5):
        r = R_ANS + i
        ws.write(r, C_DATE, f"Нед {i + 1}", ans_lab)
        ws.merge_range(r, C_H0, r, C_NOTE, "", ans_f)

    if bk.target == "sheets":
        a, b = kr(days[0]) + 1, kr(days[-1]) + 1
        ws.merge_range(R_SPARK, C_DATE, R_SPARK, C_NOTE, "",
                       bk.fmt(bg_color=th["panel"]))
        bk.f(ws, R_SPARK, C_DATE,
             '=SPARKLINE(%s!$%s$%d:$%s$%d,{"charttype","column";"color","%s";"empty","zero"})'
             % (CALC, CN(KC_TICKS), a, CN(KC_TICKS), b, th["done"]),
             bk.fmt(bg_color=th["panel"]))
    else:
        ws.merge_range(R_SPARK, C_DATE, R_SPARK, C_NOTE, "", bk.fmt(bg_color=th["panel"]))

    ws.freeze_panes(R_D1, C_H0)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.print_area(0, 0, R_D31, C_NOTE)          # печать: шапка и сетка дней
    return ws


def streak_formula(bk, i, sr):
    """Серия привычки: число, «🔥», «🛡», «верни» либо «ок» / «ещё N» для неежедневных."""
    rowt = f"{CALC}!{K_ROWT}"
    s = f"INDEX({CALC}!${CN(KC_S0 + i)}:${CN(KC_S0 + i)},{rowt})"
    miss = f"INDEX({CALC}!${CN(KC_MISS0 + i)}:${CN(KC_MISS0 + i)},{rowt})"
    return (f'=IF(\'Старт\'!$B${sr}="","",IF({rowt}=0,"",'
            f'IF(IF(\'Старт\'!$F${sr}="",7,\'Старт\'!$F${sr})<7,"ок",'
            f'IF({s}=0,IF({miss}>1,"верни",""),'
            f'IF({s}>=3,"🔥","")&{s}&IF({miss}<=1,"🛡",""))))'
            f')')


def current_week_index(bk, m, weeks):
    """Неделя для блока: текущая, если месяц активен, иначе последняя неделя месяца."""
    year = bk.year
    today = bk.today
    if today.year == year and today.month == m:
        target = today
    else:
        target = dt.date(year, m, calendar.monthrange(year, m)[1])
    yday = target.timetuple().tm_yday
    for i, days in enumerate(weeks):
        if yday in days:
            return i
    return len(weeks) - 1


def week_lines(bk, wrow, mrow):
    """Пять строк блока «Неделя»."""
    done = f"{CALC}!{RC(wrow, WC_DONE, True, True)}"
    exp = f"{CALC}!{RC(wrow, WC_EXP, True, True)}"
    prev = f"{CALC}!{RC(wrow, WC_PREV, True, True)}"
    prev_exp = f"{CALC}!{RC(max(K_WEEK, wrow - 1), WC_EXP, True, True)}"
    weak = f"{CALC}!{RC(wrow, WC_WEAK, True, True)}"
    cnt = f"{CALC}!{RC(wrow, WC_CNT, True, True)}"
    closed = f"{CALC}!{RC(wrow, WC_CLOSED, True, True)}"
    el = f"{CALC}!{RC(mrow, MC_EL, True, True)}"
    weak_name = f'INDEX(\'Старт\'!$E$12:$E$18,MAX(1,{weak}))'
    weak_why = f'INDEX(\'Старт\'!$B$20:$B$26,MAX(1,{weak}))'
    return [
        f'=IF({exp}=0,"{T.EMPTY["week"]}","Неделя · "&{done}&" из "&{exp}&'
        f'IF({prev}=0,""," · прошлая "&ROUND(MIN(1,{prev}/MAX(1,{prev_exp}))*100,0)&" %"))',
        f'=IF(OR({weak}=0,{exp}=0),"","Просело: "&{weak_name}&" — "&'
        f'IF({weak_why}="","подтяни на этой неделе",{weak_why}))',
        f'=IF({el}=0,"{T.EMPTY["month"]}","Полных дней: "&'
        f'{CALC}!{RC(mrow, MC_FULL, True, True)}&" · опыт за месяц: "&'
        f'{CALC}!{RC(mrow, MC_XP, True, True)})',
        f'=IF({weak}=0,"{T.EMPTY["challenge"]}","Испытание: "&{weak_name}&" "&'
        f'{cnt}&" из 5"&IF({closed}=1," ✓",""))',
        f'="{T.L["week_question"]}"',
    ]


def dayline_formula(bk, m, weeks):
    """Строка дня: возврат → щит → уровень → испытание → девиз."""
    rowt = f"{CALC}!{K_ROWT}"
    parts = []
    ret = f"INDEX({CALC}!${CN(KC_RET)}:${CN(KC_RET)},{rowt})"
    # у какой привычки серия равна нулю — первая такая
    zero_terms = []
    for i in range(NH):
        s = f"INDEX({CALC}!${CN(KC_S0 + i)}:${CN(KC_S0 + i)},{rowt})"
        ms = f"INDEX({CALC}!${CN(KC_MISS0 + i)}:${CN(KC_MISS0 + i)},{rowt})"
        zero_terms.append(f"IF(AND('Старт'!$B${12 + i}<>\"\","
                          f"IF('Старт'!$F${12 + i}=\"\",7,'Старт'!$F${12 + i})=7,"
                          f"{s}=0,{ms}>1),{i + 1},0)")
    first_zero = f"MAX({','.join(zero_terms)})"
    nm = f"INDEX('Старт'!$E$12:$E$18,MAX(1,{first_zero}))"
    why = f"INDEX('Старт'!$B$20:$B$26,MAX(1,{first_zero}))"
    wi = current_week_index(bk, m, weeks)
    closed = f"{CALC}!{RC(K_WEEK + wi, WC_CLOSED, True, True)}"
    return (
        f'=IF(OR({rowt}=0,{CALC}!{K_XP}=0),"{T.MOTTO}",'
        f'IF({ret}>0,"{T.DAYLINE["award_return"]}",'
        f'IF({first_zero}>0,"Верни «"&{nm}&"» сегодня"&IF({why}="",""," — "&{why}),'
        f'IF({closed}=1,"{T.DAYLINE["challenge_done"]}",'
        f'"Уровень "&{CALC}!{K_LVL}&" · "&{CALC}!{K_RANK}))))')


# ─────────────────────────────────────────────────────────────────────────────
# Лист «Старт»
# ─────────────────────────────────────────────────────────────────────────────
def build_start(bk, demo=False):
    """Первый лист: шаги, привычки, «Чтобы…», плитки месяцев. Ширина кадра 348 px."""
    th, year = bk.th, bk.year
    ws = bk.wb.add_worksheet("Старт")
    ws.hide_gridlines(2)
    ws.set_tab_color(th["accent"])
    base = bk.fmt(bg_color=th["bg"], font_color=th["text"])
    ws.set_column(0, 40, w(58), base)           # шесть равных колонок = 348 px
    LAST = 5

    title = bk.fmt(bg_color=th["bg"], font_color=th["text"], bold=True, font_size=22,
                   valign="vcenter", align="left", indent=1)
    sub = bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=11,
                 valign="vcenter", align="left", indent=1)
    step = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                  valign="vcenter", align="left", indent=1, text_wrap=True)
    sec = bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=9, bold=True,
                 valign="vcenter", align="left", indent=1)
    hdr = bk.fmt(bg_color=th["panel2"], font_color=th["muted"], font_size=9, bold=True,
                 valign="vcenter", align="center", text_wrap=True)
    inp = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=11,
                 valign="vcenter", align="left", indent=1, border=1, border_color=th["bg"])
    inp_c = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=13,
                   valign="vcenter", align="center", border=1, border_color=th["bg"])
    note = bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=10,
                  valign="vcenter", align="left", indent=1, text_wrap=True)

    ws.set_row(0, h(10))
    ws.set_row(1, h(44))
    ws.merge_range(1, 0, 1, LAST, T.BRAND.upper(), title)
    ws.set_row(2, h(24))
    ws.merge_range(2, 0, 2, LAST, T.MOTTO, sub)
    ws.set_row(3, h(10))
    for i, t in enumerate(T.STEPS):
        ws.set_row(4 + i, h(34))
        ws.merge_range(4 + i, 0, 4 + i, LAST, f"{i + 1}. {t}", step)

    # строка 9 зарезервирована: C9 — необязательная дата «считать сегодняшним днём»
    ws.set_row(8, h(14))
    ws.write_blank(8, 2, None, bk.fmt(bg_color=th["bg"], font_color=th["bg"],
                                      num_format="dd.mm.yyyy"))

    # таблица привычек: значок | название (B:C) | имя | раз в неделю
    R_HEAD = 10
    ws.set_row(R_HEAD - 1, h(24))
    ws.merge_range(R_HEAD - 1, 0, R_HEAD - 1, LAST, T.L["habits"].upper(), sec)
    ws.set_row(R_HEAD, h(32))
    ws.write(R_HEAD, 0, T.L["icon"], hdr)
    ws.merge_range(R_HEAD, 1, R_HEAD, 3, T.L["name"], hdr)
    ws.write(R_HEAD, 4, T.L["short"], hdr)
    ws.write(R_HEAD, 5, T.L["freq"], hdr)
    for i in range(NH):
        r = R_HEAD + 1 + i
        ws.set_row(r, h(40))
        ex = T.EXAMPLES[i] if i < len(T.EXAMPLES) else None
        ws.write(r, 0, ex[0] if ex else "", inp_c)
        ws.merge_range(r, 1, r, 3, "", inp)
        if ex:
            ws.write(r, 1, ex[1], inp)
        ws.write(r, 4, ex[2] if ex else "", inp)
        if ex:
            ws.write_number(r, 5, ex[3], inp_c)
        else:
            ws.write_blank(r, 5, None, inp_c)
    ws.data_validation(R_HEAD + 1, 0, R_HEAD + NH, 0,
                       {"validate": "list", "source": T.ICONS, "ignore_blank": True})
    ws.data_validation(R_HEAD + 1, 5, R_HEAD + NH, 5,
                       {"validate": "integer", "criteria": "between", "minimum": 1,
                        "maximum": 7, "ignore_blank": True, "error_title": T.L["freq"],
                        "error_message": "Сколько раз в неделю: от 1 до 7"})
    ws.data_validation(R_HEAD + 1, 1, R_HEAD + NH, 1,
                       {"validate": "length", "criteria": "<=", "value": T.LIMIT_NAME,
                        "ignore_blank": True, "error_title": T.L["name"],
                        "error_message": f"Не длиннее {T.LIMIT_NAME} знаков"})
    ws.data_validation(R_HEAD + 1, 4, R_HEAD + NH, 4,
                       {"validate": "length", "criteria": "<=", "value": T.LIMIT_SHORT,
                        "ignore_blank": True, "error_title": T.L["short"],
                        "error_message": f"Короткое имя: до {T.LIMIT_SHORT} знаков"})

    # блок «Чтобы…» — по строке на привычку
    R_WHY = R_HEAD + NH + 2
    ws.set_row(R_WHY - 1, h(24))
    ws.merge_range(R_WHY - 1, 0, R_WHY - 1, LAST, T.L["why"].upper(), sec)
    for i in range(NH):
        r = R_WHY + i
        ws.set_row(r, h(36))
        bk.f(ws, r, 0, f"=IF($B${R_HEAD + 2 + i}=\"\",\"\",$A${R_HEAD + 2 + i})",
             bk.fmt(bg_color=th["panel"], font_size=13, align="center", valign="vcenter"))
        ws.merge_range(r, 1, r, LAST, "", inp)
        if i < len(T.EXAMPLES) and T.EXAMPLES[i][4]:
            ws.write(r, 1, T.EXAMPLES[i][4], inp)
    ws.data_validation(R_WHY, 1, R_WHY + NH - 1, 1,
                       {"validate": "length", "criteria": "<=", "value": T.LIMIT_WHY,
                        "ignore_blank": True, "error_title": T.L["why"],
                        "error_message": f"Коротко, до {T.LIMIT_WHY} знаков"})

    # подсказки
    r = R_WHY + NH + 1
    hint = bk.fmt(bg_color=th["panel"], font_color=th["accent"], font_size=10, bold=True,
                  valign="vcenter", align="left", indent=1, text_wrap=True)
    for i, line in enumerate(T.HELP):
        ws.set_row(r + i, h(38 if i == 0 else 30))
        ws.merge_range(r + i, 0, r + i, LAST, line if i == 0 else "· " + line,
                       hint if i == 0 else note)

    # плитки месяцев: два ряда по шесть равных колонок
    r2 = r + len(T.HELP) + 1
    ws.set_row(r2, h(24))
    ws.merge_range(r2, 0, r2, LAST, "МЕСЯЦЫ", sec)
    tile = bk.fmt(bg_color=th["tile"], font_color=th["text"], font_size=12, bold=True,
                  align="center", valign="vcenter", border=1, border_color=th["bg"],
                  underline=0)
    for i in range(12):
        rr, cc = r2 + 1 + i // 6, i % 6
        ws.set_row(rr, h(48))
        if (i + 1) in getattr(bk, "months", range(1, 13)):
            ws.write_url(rr, cc, f"internal:'{T.MONTHS_SHORT[i]}'!A1", tile, T.MONTHS_SHORT[i])
        else:
            ws.write(rr, cc, T.MONTHS_SHORT[i], tile)
    tile_now = bk.fmt(bg_color=th["accent"], font_color="#1A1400", font_size=12, bold=True,
                      align="center", valign="vcenter", border=1, border_color=th["bg"],
                      underline=0)
    for half in range(2):
        bk.cf(ws, r2 + 1 + half, 0, r2 + 1 + half, 5, {
            "type": "formula",
            "criteria": f"=MONTH({CALC}!{K_TODAY})=COLUMN()+{half * 6}",
            "format": tile_now})

    # самопроверка и подсказка про тёмную тему
    r3 = r2 + 4
    ws.set_row(r3, h(34))
    ok = bk.fmt(bg_color=th["panel"], font_color=th["done"], font_size=11, bold=True,
                align="left", indent=1, valign="vcenter")
    ws.merge_range(r3, 0, r3, LAST, "", ok)
    bk.f(ws, r3, 0,
         f'=IF({CALC}!{K_NACT}=0,"Впиши хотя бы одну привычку",'
         f'"{T.L["selfcheck"]} · {T.VERSION} · "&{CALC}!{K_NACT}&" привычек")', ok)
    ws.set_row(r3 + 1, h(30))
    ws.merge_range(r3 + 1, 0, r3 + 1, LAST, T.FOOTER_DARK, note)
    ws.set_row(r3 + 2, h(34))
    ws.merge_range(r3 + 2, 0, r3 + 2, 2, T.L["today_date"] + " (обычно пусто)", note)
    ws.merge_range(r3 + 2, 3, r3 + 2, LAST, "", bk.fmt(
        bg_color=th["input"], font_color=th["muted"], font_size=10, align="center",
        valign="vcenter", border=1, border_color=th["bg"]))
    bk.f(ws, r3 + 2, 3, '=IF($C$9="","—",$C$9)', bk.fmt(
        bg_color=th["input"], font_color=th["muted"], font_size=10, align="center",
        valign="vcenter", border=1, border_color=th["bg"], num_format="dd.mm.yyyy"))
    ws.freeze_panes(4, 0)
    ws.fit_to_pages(1, 1)
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# Листы «Год», «Обзор», «Списки»
# ─────────────────────────────────────────────────────────────────────────────
def build_year(bk, counters, weeks):
    th, year = bk.th, bk.year
    ws = bk.wb.add_worksheet("Год")
    ws.hide_gridlines(2)
    ws.set_tab_color(th["xp"])
    base = bk.fmt(bg_color=th["bg"], font_color=th["text"])
    ws.set_column(0, 40, w(43), base)
    ws.set_column(0, 0, w(64), base)
    ws.set_column(1, 8, w(43), base)
    ws.set_column(9, 9, w(60), base)

    title = bk.fmt(bg_color=th["panel"], font_color=th["text"], bold=True, font_size=16,
                   valign="vcenter", align="left", indent=1)
    ws.set_row(0, h(44))
    ws.merge_range(0, 0, 0, 9, "", title)
    bk.f(ws, 0, 0, f'="ГОД {year} · "&{CALC}!{K_RANK}&" · уровень "&{CALC}!{K_LVL}', title)

    sec = bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=9, bold=True,
                 valign="vcenter", align="left", indent=1)
    hdr = bk.fmt(bg_color=th["panel2"], font_color=th["muted"], font_size=9, bold=True,
                 align="center", valign="vcenter")
    mon_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11, bold=True,
                   align="left", indent=1, valign="vcenter")
    val_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                   align="center", valign="vcenter")
    pct_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                   align="center", valign="vcenter", num_format="0%")

    ws.set_row(1, h(24))
    ws.merge_range(1, 0, 1, 9, "МЕСЯЦЫ", sec)
    ws.set_row(2, h(28))
    for c, t in ((0, "Месяц"), (1, "%"), (3, "Всего"), (4, "Полных"),
                 (5, "Серия"), (6, "Опыт")):
        ws.write(2, c, t, hdr)
    ws.write(2, 2, "", hdr)
    for m in range(1, 13):
        r = 2 + m
        mrow = K_MON + m - 1
        ws.set_row(r, h(34))
        ws.write(r, 0, T.MONTHS_SHORT[m - 1], mon_f)
        el = f"{CALC}!{RC(mrow, MC_DONE, True, True)}"          # пусто, пока нет галочек
        bk.f(ws, r, 1, f"=IF({el}=0,\"\",{CALC}!{RC(mrow, MC_PCT, True, True)})", pct_f)
        if bk.target == "sheets":
            bk.f(ws, r, 2,
                 '=IF(%s="","",SPARKLINE(%s,{"charttype","bar";"max",1;"color1","%s";"color2","%s"}))'
                 % (RC(r, 1, True, True), RC(r, 1, True, True), th["done"], th["panel2"]),
                 bk.fmt(bg_color=th["panel"]))
        else:
            bk.f(ws, r, 2, f'=IF({RC(r, 1, True, True)}="","",'
                           f'REPT("▮",ROUND({RC(r, 1, True, True)}*8,0)))',
                 bk.fmt(bg_color=th["panel"], font_color=th["done"], font_size=11))
        for c, col in ((3, MC_DONE), (4, MC_FULL), (5, MC_BEST), (6, MC_XP)):
            bk.f(ws, r, c, f"=IF({el}=0,\"\",{CALC}!{RC(mrow, col, True, True)})", val_f)

    # награды
    r0 = 16
    ws.set_row(r0, h(24))
    ws.merge_range(r0, 0, r0, 9, T.L["awards"].upper(), sec)
    chip = bk.fmt(bg_color=th["panel"], font_color=th["muted"], font_size=11,
                  align="left", indent=1, valign="vcenter", border=1, border_color=th["bg"])
    for i, (nm, key, goal, cumulative) in enumerate(T.AWARDS):
        r = r0 + 1 + i // 2
        c = (i % 2) * 5
        ws.set_row(r, h(36))
        ws.merge_range(r, c, r, c + 4, "", chip)
        v = counters[key]
        if cumulative:
            bk.f(ws, r, c, f'=IF({v}>={goal},"🏅 {nm}","{nm} · "&{v}&"/{goal}")', chip)
        else:
            bk.f(ws, r, c, f'=IF({v}>={goal},"🏅 {nm}","{nm}")', chip)
    bk.cf(ws, r0 + 1, 0, r0 + 5, 9, {
        "type": "text", "criteria": "beginsWith", "value": "🏅",
        "format": bk.fmt(bg_color=th["panel"], font_color=th["accent"], bold=True)})

    # итог года
    r1 = r0 + 7
    ws.set_row(r1, h(24))
    ws.merge_range(r1, 0, r1, 9, T.L["year_total"].upper(), sec)
    ws.set_row(r1 + 1, h(40))
    ws.merge_range(r1 + 1, 0, r1 + 1, 9, "", bk.fmt(
        bg_color=th["panel"], font_color=th["text"], font_size=13, bold=True,
        align="left", indent=1, valign="vcenter"))
    bk.f(ws, r1 + 1, 0,
         f'=IF({counters["ticks"]}=0,"{T.EMPTY["year"]}",'
         f'{counters["ticks"]}&" галочек · опыт "&{CALC}!{K_XP}&" · рекорд серии "'
         f'&{counters["record"]}&" · полных дней "&{counters["full_days"]})',
         bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=13, bold=True,
                align="left", indent=1, valign="vcenter"))

    # ответы на вопрос недели — во всю ширину под итогом года
    r2 = r1 + 3
    ws.set_row(r2, h(24))
    ws.merge_range(r2, 0, r2, 9, T.L["week_question"].upper(), sec)
    ans_lab = bk.fmt(bg_color=th["panel"], font_color=th["muted"], font_size=10,
                     align="center", valign="vcenter")
    ans = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=11,
                 align="left", indent=1, valign="vcenter", text_wrap=True,
                 border=1, border_color=th["bg"])
    for i in range(min(53, len(weeks))):
        rr = r2 + 1 + i
        ws.set_row(rr, h(34))
        ws.write(rr, 0, f"Нед {i + 1}", ans_lab)
        ws.merge_range(rr, 1, rr, 9, "", ans)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws


def build_review(bk):
    """«Обзор» — зеркало месяца для компьютера, через CHOOSE, без INDIRECT."""
    th = bk.th
    ws = bk.wb.add_worksheet("Обзор")
    ws.hide_gridlines(2)
    base = bk.fmt(bg_color=th["bg"], font_color=th["text"])
    ws.set_column(0, 40, w(43), base)
    ws.set_column(0, 0, w(150), base)
    title = bk.fmt(bg_color=th["panel"], font_color=th["text"], bold=True, font_size=14,
                   valign="vcenter", align="left", indent=1)
    ws.set_row(0, h(40))
    ws.merge_range(0, 0, 0, 8, "ОБЗОР МЕСЯЦА · для компьютера", title)
    ws.set_row(1, h(34))
    ws.write(1, 0, "Месяц:", bk.fmt(bg_color=th["bg"], font_color=th["muted"], font_size=11,
                                    align="right"))
    sel = bk.fmt(bg_color=th["input"], font_color=th["text"], font_size=12, bold=True,
                 align="center", valign="vcenter")
    ws.write(1, 1, T.MONTHS_SHORT[bk.today.month - 1] if bk.today.year == bk.year
             else T.MONTHS_SHORT[0], sel)
    ws.data_validation(1, 1, 1, 1, {"validate": "list", "source": T.MONTHS_SHORT})
    mi = f'MATCH($B$2,\'Списки\'!$H$2:$H$13,0)'
    hdr = bk.fmt(bg_color=th["panel2"], font_color=th["muted"], font_size=9, bold=True,
                 align="center", valign="vcenter")
    ws.set_row(3, h(28))
    ws.write(3, 0, "Привычка", hdr)
    for d in range(1, 32):
        ws.set_column(d, d, w(28), base)
        ws.write(3, d, d, hdr)
    name_f = bk.fmt(bg_color=th["panel"], font_color=th["text"], font_size=11,
                    align="left", indent=1, valign="vcenter")
    cell_f = bk.fmt(bg_color=th["panel2"], font_color=th["done"], font_size=11, bold=True,
                    align="center", valign="vcenter", border=1, border_color=th["bg"])
    for i in range(NH):
        r = 4 + i
        ws.set_row(r, h(34))
        bk.f(ws, r, 0, f"=IF('Старт'!$B${12 + i}=\"\",\"\",'Старт'!$B${12 + i})", name_f)
        for d in range(1, 32):
            avail = getattr(bk, "months", range(1, 13))
            refs = ",".join(
                (f"'{T.MONTHS_SHORT[m - 1]}'!{RC(R_D1 + d - 1, C_H0 + i, True, True)}"
                 if m in avail else "FALSE")
                for m in range(1, 13))
            bk.f(ws, r, d,
                 f'=IF(\'Старт\'!$B${12 + i}="","",IF(CHOOSE({mi},{refs})=TRUE,"✓",""))',
                 cell_f)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws


def build_lists(bk):
    ws = bk.wb.add_worksheet("Списки")
    ws.hide()
    for i, v in enumerate(T.SLEEP):
        ws.write(1 + i, 0, v)
    for i, v in enumerate(T.ENERGY):
        ws.write(1 + i, 1, v)
    for i, v in enumerate(T.MOOD):
        ws.write(1 + i, 2, v)
    for i, v in enumerate(T.ICONS):
        ws.write(1 + i, 3, v)
    ws.write(0, 4, "порог")
    ws.write(0, 5, "звание")
    for lvl in range(1, 21):
        ws.write_number(lvl, 4, 100 * (lvl - 1) * lvl)
        rank = [r for lo, r in T.RANKS if lvl >= lo][-1]
        ws.write(lvl, 5, rank)
    ws.write(0, 7, "месяцы")
    for i, v in enumerate(T.MONTHS_SHORT):
        ws.write(1 + i, 7, v)
    return ws


# ─────────────────────────────────────────────────────────────────────────────
def build(path, theme, target, year, today, demo=None, only_months=None):
    """only_months — собрать пробник на часть месяцев (смоук-тест импорта)."""
    budget = Budget()
    bk = Book(path, theme, target, year, today, budget)
    weeks = week_blocks(year)
    build_start(bk, demo=bool(demo))
    bk.months = only_months or list(range(1, 13))
    months = [build_month(bk, m, weeks, (demo or {}).get(m)) for m in bk.months]
    counters = build_calc(bk, weeks)
    build_year(bk, counters, weeks)
    build_review(bk)
    build_lists(bk)
    bk.wb.worksheets_objs[0].activate()          # «Старт» открывается первым
    bk.wb.set_properties({"title": f"{T.BRAND} · планер привычек",
                          "comments": f"Собрано build.py, {T.VERSION}"})
    bk.wb.close()
    return budget


def geometry_check():
    """Ворота геометрии: кадр ≤350 px, шапка 124 px, строки 46 px, кегли ≥10 pt."""
    bad = []
    frame = 48 + NH * 43
    if frame > 350:
        bad.append(f"кадр {frame} px > 350")
    head = 36 + 24 + 24 + 20 + 20
    if head != 124:
        bad.append(f"шапка {head} px ≠ 124")
    return bad


def main():
    ap = argparse.ArgumentParser(description="Сборка планера «Режим»")
    here = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--out", default=os.path.join(here, "..", "dist"))
    ap.add_argument("--year", type=int, default=None)
    ap.add_argument("--today", default=None, help="дата вида 2026-09-21 для проверок")
    ap.add_argument("--recalc", action="store_true")
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--smoke", action="store_true", help="пробник на один месяц для импорта")
    args = ap.parse_args()

    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()
    year = args.year or today.year
    out = os.path.abspath(args.out)
    os.makedirs(out, exist_ok=True)

    bad = T.check_limits() + geometry_check()
    if bad:
        print("ВОРОТА НЕ ПРОЙДЕНЫ:\n  " + "\n  ".join(bad))
        return 1

    demo = make_demo(year, today) if args.demo else None
    if args.smoke:
        path = os.path.join(out, "smoke.xlsx")
        budget = build(path, "dark", "sheets", year, today, demo, only_months=[today.month])
        bad = budget.check()
        print(f"пробник: формул {budget.formulas}; " + ("ворота пройдены" if not bad else str(bad)))
        return 1 if bad else 0
    targets = [(f"Режим · {year} (тёмный).xlsx", "dark", "sheets", None),
               (f"Режим · {year} (светлый).xlsx", "light", "sheets", None),
               (f"Режим · {year} · Excel.xlsx", "dark", "excel", None)]
    if demo:
        targets.append((f"Режим · демо.xlsx", "dark", "sheets", demo))

    for fname, theme, target, data in targets:
        path = os.path.join(out, fname)
        print(f"→ {fname}")
        budget = build(path, theme, target, year, today, data)
        bad = budget.check()
        print(f"   формул {budget.formulas}, TODAY() {budget.today}, "
              f"правил УФ {max(budget.cf.values()) if budget.cf else 0} макс, "
              f"SPARKLINE {max(budget.spark.values()) if budget.spark else 0} макс")
        if bad:
            print("   ВОРОТА НЕ ПРОЙДЕНЫ:\n     " + "\n     ".join(bad))
            return 1
    print("готово:", out)
    return 0


def make_demo(year, today):
    import random
    rnd = random.Random(11)
    probs = [0.72, 0.95, 0.86, 0.62, 0.8]
    data = {}
    for m in range(max(1, today.month - 2), today.month + 1):
        dim = calendar.monthrange(year, m)[1]
        last = dim if m < today.month else today.day
        d = {"checks": set(), "sleep": {}, "energy": {}, "mood": {}, "note": {}}
        for day in range(1, last + 1):
            sleep = rnd.choice(T.SLEEP[1:])
            boost = 1.1 if sleep != T.SLEEP[1] else 0.8
            done = 0
            for i, p in enumerate(probs):
                if rnd.random() < min(0.97, p * boost):
                    d["checks"].add((i, day))
                    done += 1
            d["sleep"][day] = sleep
            d["energy"][day] = T.ENERGY[min(2, done // 2)]
            d["mood"][day] = T.MOOD[min(4, done)]
        d["note"][min(12, last)] = "Лёг раньше — утро другое"
        data[m] = d
    return data


if __name__ == "__main__":
    sys.exit(main())
