# -*- coding: utf-8 -*-
"""Штатное расписание ООО «МИР» (ШР-2026) — заполненная версия.

Источники:
  - «АУДИТ_СОУТ/ШР  2026.xlsx» — исходная таблица (подразделения, должности,
    штатные единицы; в колонке окладов стояли пометки занятости, а не суммы);
  - «05_Приказ_штатное_расписание.docx» — приказ № 11 об утверждении ШР-2026
    (44 штатные единицы, оклады не ниже 32 511,60 руб.);
  - черновик СОУТ (список 25 работников на 11.02.2026) — сверка занятости.

Заполнено: наименование организации, утверждение приказом № 11, период,
итого по листу/документу (44 — пересчитано и сверено с приказом),
руководитель кадровой службы (Специалист по кадрам / Максимова Л.В.).
Оставлено пустым (данных в репозитории нет): номер/дата документа,
оклады ([оклад]), ФИО главного бухгалтера, подписи.

Запуск: python build_shr.py  (выход — ШР_2026_ЗАПОЛНЕННОЕ.docx в этой папке)
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/home/user/MIR/АУДИТ_СОУТ/ШР_2026_ЗАПОЛНЕННОЕ.docx"

ORG_FULL = "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ «МИР»"
ORG_SHORT = "ООО «МИР»"
REQ = "ОГРН 1156196053881 · ИНН 6166094734 · КПП 616601001"
ADR = "344093, РФ, Ростовская область, г. Ростов-на-Дону, ул. Туполева, зд. 16Е, оф. 4.13"

# (подразделение, код, должность, кол-во) — ровно как в исходном xlsx
ROWS = [
    ("Администрация", "1", "Генеральный директор", 1),
    ("Администрация", "1", "Главный бухгалтер", 1),
    ("Администрация", "1", "Бухгалтер", 4),
    ("Администрация", "1", "Бухгалтер-кассир", 2),
    ("Администрация", "1", "Руководитель группы документационного сопровождения", 1),
    ("Администрация", "1", "Менеджер по IT-проектам", 1),
    ("Администрация", "1", "Менеджер", 2),
    ("Администрация", "1", "Делопроизводитель", 1),
    ("Администрация", "1", "Специалист по кадрам", 1),
    ("Администрация", "1", "Диспетчер", 1),
    ("ОП ООО «МИР»", "1", "Начальник производства", 1),
    ("ОП ООО «МИР»", "1", "Заведующий складом", 1),
    ("ОП ООО «МИР»", "1", "Прессовщик", 2),
    ("ОП ООО «МИР»", "1", "Сортировщик", 3),
    ("ОП ООО «МИР»", "1", "Водитель", 10),
    ("ОП ООО «МИР»", "1", "Грузчик", 6),
    ("ОП ООО «МИР»", "1", "Рабочий охраны", 2),
    ("ОП ООО «МИР»", "1", "Механик", 1),
    ("ОП ООО «МИР»", "1", "Приемщик", 1),
    ("ОП ООО «МИР»", "1", "Водитель погрузчика", 2),
]
TOTAL = sum(r[3] for r in ROWS if r[3] is not None)
assert TOTAL == 44, f"итого {TOTAL}, а приказ № 11 — 44"

NOTE = (
    "Примечания: "
    "1) оклады устанавливаются штатным расписанием и закрепляются в трудовых договорах "
    "(размеры в настоящем документе не указаны, минимальный оклад — 32 511,60 руб. по п. 3 "
    "приказа № 11); "
    "2) фактическое замещение штатных единиц: водитель — 8 из 10 (автомобилей — 10), "
    "водитель погрузчика — 1 из 2; штатные единицы не замещены: главный бухгалтер, "
    "заведующий складом, прессовщик — 2, сортировщик — 3, приемщик; "
    "3) удаленная работа: менеджер по IT-проектам, менеджер — 2, диспетчер."
)


def _run(p, text, bold=False, size=10, italic=False):
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.name = "Times New Roman"
    r.font.size = Pt(size)
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Times New Roman")
    return r


def _par(doc, text="", bold=False, size=10, align=None, italic=False,
         space_after=None, space_before=None):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    if space_after is not None:
        pf.space_after = Pt(space_after)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    if text:
        _run(p, text, bold=bold, size=size, italic=italic)
    return p


def cell_borders(cell, sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    old = tcPr.find(qn("w:tcBorders"))
    if old is not None:
        tcPr.remove(old)
    b = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), sz)
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        b.append(e)
    tcPr.append(b)


def write_cell(cell, text, bold=False, size=9, align=None,
               vertical=WD_ALIGN_VERTICAL.CENTER):
    cell.text = ""
    for i, ln in enumerate(text.split("\n")):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        _run(p, ln, bold=bold, size=size)
    cell.vertical_alignment = vertical
    cell_borders(cell)


def set_widths(table, widths_cm, target_twips=None):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for tag in ("w:tblLayout", "w:tblW", "w:tblInd"):
        old = tblPr.find(qn(tag))
        if old is not None:
            tblPr.remove(old)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    tw = [int(round(w / 2.54 * 1440)) for w in widths_cm]
    target = target_twips if target_twips else int(round(sum(widths_cm) / 2.54 * 1440))
    tw[-1] += target - sum(tw)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(sum(tw)))
    w.set(qn("w:type"), "dxa")
    tblPr.append(w)
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), "0")
    ind.set(qn("w:type"), "dxa")
    tblPr.append(ind)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for x in tw:
        g = OxmlElement("w:gridCol")
        g.set(qn("w:w"), str(x))
        grid.append(g)
    tbl.insert(list(tbl).index(tblPr) + 1, grid)
    for row in table.rows:
        for tc in row._tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            old = tcPr.find(qn("w:tcW"))
            if old is not None:
                tcPr.remove(old)
    for ri, row in enumerate(table.rows):
        col = 0
        for tc in row._tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            gs = tcPr.find(qn("w:gridSpan"))
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            width = sum(tw[col:col + span])
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(width))
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)
            col += span


def build():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(10)
    st.paragraph_format.space_after = Pt(4)
    rPr = st.element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), "Times New Roman")

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin, sec.right_margin = Cm(1.8), Cm(1.5)   # полоса 17,7 см = сумме колонок
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    # точная ширина полосы в twips так, как её сохранит python-docx (round(emu/635))
    usable_twips = (int(round(int(sec.page_width) / 635))
                    - int(round(int(sec.left_margin) / 635))
                    - int(round(int(sec.right_margin) / 635)))

    # шапка
    _par(doc, ORG_FULL, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    _par(doc, REQ, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0)
    _par(doc, ADR, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    _par(doc, "ШТАТНОЕ РАСПИСАНИЕ (ШР-2026)", bold=True, size=12,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    _par(doc, f"утверждено приказом {ORG_SHORT} № 11 от «___» __________ 2026 г.",
         size=9, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    _par(doc, "Номер документа: ______________        Дата составления: «___» __________ 2026 г.",
         size=10, space_after=2)
    _par(doc, "На период: с 14 января 2026 г.", size=10, space_after=8)

    # таблица
    W = [3.3, 0.9, 6.5, 1.4, 5.6]
    body = [r for r in ROWS if r[1] is not None]
    t = doc.add_table(rows=2, cols=5)
    hdr = ["Структурное подразделение", "код",
           "Должность (специальность, профессия),\nразряд, класс (категория) квалификации",
           "Количество\nштатных\nединиц", "Тарифная ставка\n(оклад) и пр.,\nруб."]
    for i, h in enumerate(hdr):
        c = t.rows[0].cells[i]
        write_cell(c, h, bold=True, size=8,
                   align=WD_ALIGN_PARAGRAPH.CENTER if i != 2 else WD_ALIGN_PARAGRAPH.LEFT)
    for i in range(5):
        write_cell(t.rows[1].cells[i], str(i + 1), size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

    last_podr = None
    for podr, kod, dolg, kolichestvo in body:
        if last_podr is not None and podr != last_podr:
            # тонкая разделительная строка между блоками подразделений (как в исходнике)
            sep = t.add_row()
            sep.cells[0].merge(sep.cells[4])
            write_cell(sep.cells[0], " ", size=2)
            sep.height = Pt(4)
        row = t.add_row()
        first = podr if podr != last_podr else ""
        last_podr = podr
        write_cell(row.cells[0], first, size=9, vertical=WD_ALIGN_VERTICAL.TOP)
        write_cell(row.cells[1], kod, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(row.cells[2], dolg, size=9, vertical=WD_ALIGN_VERTICAL.TOP)
        write_cell(row.cells[3], str(kolichestvo), size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        write_cell(row.cells[4], "[оклад]", size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    for label in ("Итого по листу", "Итого по документу"):
        row = t.add_row()
        c = row.cells[0]
        for cc in row.cells[1:4]:
            c = c.merge(cc)
        write_cell(c, label, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.LEFT)
        write_cell(row.cells[4], str(TOTAL), bold=True, size=9,
                   align=WD_ALIGN_PARAGRAPH.CENTER)

    set_widths(t, W, usable_twips)

    _par(doc, "", space_after=2)
    _par(doc, NOTE, size=9, space_after=14)

    # подписи
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    _run(p, "Руководитель кадровой службы", size=10)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    _run(p, "Специалист по кадрам                    _______________  / Максимова Л.В. /", size=10)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    _run(p, "Главный бухгалтер                       _______________  /                  /", size=10)

    doc.save(OUT)
    print("SAVED:", OUT)
    print("Итого штатных единиц:", TOTAL)


if __name__ == "__main__":
    build()
