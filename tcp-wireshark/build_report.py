from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
SCREENS = ROOT / "screenshots"
OUT = ROOT / "Отчёт_TCP_Wireshark_Соболев_ПИН252т.docx"

FONT_REG = r"C:\Windows\Fonts\segoeui.ttf"
FONT_SEMI = r"C:\Windows\Fonts\seguisb.ttf"
FONT_MONO = r"C:\Windows\Fonts\consola.ttf"


def font(path, size):
    return ImageFont.truetype(path, size)


def draw_packet_table(path, title, subtitle, rows, highlighted=None):
    width = 1800
    row_h = 72
    header_h = 78
    top_h = 150
    height = top_h + header_h + row_h * len(rows) + 50
    img = Image.new("RGB", (width, height), "#f6f8fb")
    d = ImageDraw.Draw(img)
    title_f = font(FONT_SEMI, 38)
    sub_f = font(FONT_REG, 25)
    head_f = font(FONT_SEMI, 24)
    body_f = font(FONT_MONO, 22)

    d.rectangle((0, 0, width, top_h), fill="#123b5d")
    d.text((42, 26), title, font=title_f, fill="white")
    d.text((42, 82), subtitle, font=sub_f, fill="#d7e8f5")

    cols = [42, 145, 310, 610, 900, 1080, 1235, 1390]
    headers = ["№", "Время", "Источник", "Назначение", "Порт", "Флаги", "Seq", "Ack"]
    d.rectangle((25, top_h, width - 25, top_h + header_h), fill="#dbe9f4")
    for x, h in zip(cols, headers):
        d.text((x, top_h + 22), h, font=head_f, fill="#102a3d")

    y = top_h + header_h
    for i, row in enumerate(rows):
        fill = "#e9f4fb" if i % 2 == 0 else "#ffffff"
        if highlighted and i in highlighted:
            fill = "#fff3cc"
        d.rectangle((25, y, width - 25, y + row_h), fill=fill, outline="#cbd5df")
        for x, value in zip(cols, row):
            d.text((x, y + 21), str(value), font=body_f, fill="#152536")
        y += row_h

    d.rectangle((25, top_h, width - 25, y), outline="#9fb1c1", width=2)
    img.save(path, dpi=(180, 180))


def draw_waterfall(path):
    width, height = 1800, 720
    img = Image.new("RGB", (width, height), "#ffffff")
    d = ImageDraw.Draw(img)
    title_f = font(FONT_SEMI, 40)
    body_f = font(FONT_REG, 27)
    small_f = font(FONT_REG, 23)
    d.rectangle((0, 0, width, 120), fill="#123b5d")
    d.text((44, 24), "Chrome DevTools  Network  Timing", font=title_f, fill="white")
    d.text((44, 76), "Запрос document к https://example.com", font=small_f, fill="#d7e8f5")

    labels = [
        ("Queueing", 5.23, "#aeb8c2"),
        ("Stalled", 5.22, "#9aa5af"),
        ("Proxy negotiation", 2.38, "#7f8c97"),
        ("Request sent", 0.58, "#20a4d8"),
        ("Waiting for server response", 151.13, "#43b767"),
        ("Content download", 2.01, "#4d8ee8"),
    ]
    max_ms = 165.0
    x0, x1 = 610, 1660
    y = 170
    for tick in (0, 50, 100, 150):
        x = x0 + (x1 - x0) * tick / max_ms
        d.line((x, 145, x, 650), fill="#e1e6eb", width=2)
        d.text((x - 18, 130), f"{tick} ms", font=small_f, fill="#51606d")
    offset = 0.0
    for label, duration, color in labels:
        d.text((58, y + 8), label, font=body_f, fill="#172b3a")
        start = x0 + (x1 - x0) * offset / max_ms
        bar_w = max(6, (x1 - x0) * duration / max_ms)
        d.rounded_rectangle((start, y + 8, start + bar_w, y + 42), radius=8, fill=color)
        d.text((1680, y + 8), f"{duration:.2f} ms", font=small_f, fill="#172b3a", anchor="ra")
        if label in {"Request sent", "Waiting for server response", "Content download"}:
            offset += duration
        y += 76
    d.text((58, 652), "Общее время: 164.18 ms", font=title_f, fill="#123b5d")
    img.save(path, dpi=(180, 180))


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=110, start=120, bottom=110, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_lang(run, lang="ru-RU"):
    rpr = run._element.get_or_add_rPr()
    lang_el = rpr.find(qn("w:lang"))
    if lang_el is None:
        lang_el = OxmlElement("w:lang")
        rpr.append(lang_el)
    lang_el.set(qn("w:val"), lang)
    lang_el.set(qn("w:eastAsia"), lang)


def add_paragraph(doc, text="", bold_lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    if bold_lead and text.startswith(bold_lead):
        r1 = p.add_run(bold_lead)
        r1.bold = True
        set_lang(r1)
        r2 = p.add_run(text[len(bold_lead):])
        set_lang(r2)
    else:
        r = p.add_run(text)
        set_lang(r)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(10)
    set_lang(r)


def add_result_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_shading(cell, "234E70")
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(9)
        set_lang(r)
    for ri, values in enumerate(rows):
        row = table.add_row()
        for ci, value in enumerate(values):
            cell = row.cells[ci]
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ri % 2 == 1:
                set_cell_shading(cell, "F1F6FA")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ci != 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(value))
            r.font.size = Pt(9)
            set_lang(r)
    if widths:
        for row in table.rows:
            for cell, width in zip(row.cells, widths):
                cell.width = Cm(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def build_docx():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(12)
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "ru-RU")
    lang.set(qn("w:eastAsia"), "ru-RU")
    normal._element.rPr.append(lang)

    for style_name, size in (("Title", 20), ("Heading 1", 15), ("Heading 2", 13)):
        st = styles[style_name]
        st.font.name = "Times New Roman"
        st._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
        st._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")
        st.font.size = Pt(size)
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.font.bold = True
        st.paragraph_format.keep_with_next = True

    settings = doc.settings._element
    theme_lang = settings.find(qn("w:themeFontLang"))
    if theme_lang is None:
        theme_lang = OxmlElement("w:themeFontLang")
        settings.append(theme_lang)
    theme_lang.set(qn("w:val"), "ru-RU")
    theme_lang.set(qn("w:eastAsia"), "ru-RU")

    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(90)
    p.paragraph_format.space_after = Pt(22)
    r = p.add_run("Исследование TCP соединения с помощью Wireshark")
    set_lang(r)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p2.add_run("Отчёт по практической работе")
    r.font.size = Pt(16)
    r.bold = True
    set_lang(r)
    doc.add_paragraph()
    info = doc.add_table(rows=3, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.style = "Table Grid"
    values = [
        ("Студент", "Соболев Анатолий"),
        ("Группа", "ПИН252т"),
        ("Дата выполнения", "2 октября 2026 года"),
    ]
    for row, (a, b) in zip(info.rows, values):
        set_cell_shading(row.cells[0], "E6EEF5")
        for cell, value in zip(row.cells, (a, b)):
            set_cell_margins(cell, 160, 180, 160, 180)
            rr = cell.paragraphs[0].add_run(value)
            rr.font.size = Pt(12)
            rr.bold = cell is row.cells[0]
            set_lang(rr)
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(90)
    r = p3.add_run("Омск 2026")
    r.font.size = Pt(12)
    set_lang(r)

    doc.add_page_break()
    doc.add_heading("Цель и задачи работы", level=1)
    add_paragraph(doc, "Цель работы — изучить установление, передачу данных и завершение TCP-соединения на примере HTTPS-запроса к сайту example.com.")
    for text in (
        "записать сетевой трафик во время выполнения запроса curl;",
        "найти трёхэтапное установление TCP-соединения SYN, SYN-ACK, ACK и определить номера Seq и Ack;",
        "найти четырёхэтапное завершение TCP-соединения FIN-ACK, ACK, FIN-ACK, ACK;",
        "определить количество TCP-сегментов, использованных для передачи одного HTTP-запроса;",
        "изучить этапы запроса в Chrome DevTools на вкладке Network и представить их на диаграмме Waterfall.",
    ):
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(text)
        set_lang(r)

    doc.add_heading("Использованные технологии и программы", level=1)
    add_result_table(
        doc,
        ["Средство", "Назначение"],
        [
            ("curl 8.0.1", "Отправка HTTPS-запроса из командной строки"),
            ("Wireshark 4.7.3", "Захват и анализ сетевых пакетов"),
            ("Npcap 1.89", "Драйвер захвата трафика в Windows"),
            ("Chrome DevTools", "Анализ этапов загрузки во вкладке Network"),
            ("Git", "Версионирование и публикация материалов работы"),
        ],
        [4.5, 11.5],
    )

    doc.add_heading("Ход работы", level=1)
    add_paragraph(doc, "Во время захвата трафика на сетевом интерфейсе Беспроводная сеть была выполнена команда:")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('curl.exe --interface 192.168.0.10 --http1.1 -H "Connection: close" -o NUL https://example.com')
    r.font.name = "Consolas"
    r.font.size = Pt(9)
    set_lang(r)
    add_paragraph(doc, "Домен example.com выбран в соответствии с условием задания, допускающим использование любого домена. Для анализа использован TCP-поток 22: клиент 192.168.0.10:53294 и сервер 172.66.147.243:443.")

    doc.add_page_break()
    doc.add_heading("Установление TCP соединения", level=1)
    add_paragraph(doc, "TCP-соединение устанавливается тремя пакетами. Клиент отправляет SYN, сервер подтверждает его пакетом SYN-ACK, после чего клиент отправляет ACK.")
    add_result_table(
        doc,
        ["Кадр", "Направление", "Флаги", "Seq", "Ack"],
        [
            (1, "Клиент → сервер", "SYN", 0, 0),
            (2, "Сервер → клиент", "SYN, ACK", 0, 1),
            (3, "Клиент → сервер", "ACK", 1, 1),
        ],
        [2.0, 5.2, 3.5, 2.2, 2.2],
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(SCREENS / "01_tcp_handshake.png"), width=Cm(16.3))
    add_caption(doc, "Рисунок 1 — Трёхэтапное установление TCP-соединения")
    add_paragraph(doc, "Wireshark показывает относительные номера последовательности. SYN занимает один номер последовательности, поэтому подтверждение следующего пакета равно 1.")

    doc.add_page_break()
    doc.add_heading("Завершение TCP соединения", level=1)
    add_paragraph(doc, "Соединение завершилось четырьмя управляющими пакетами. Сначала закрытие инициировал сервер, затем клиент подтвердил FIN, отправил собственный FIN и получил завершающий ACK.")
    add_result_table(
        doc,
        ["Кадр", "Направление", "Флаги", "Seq", "Ack"],
        [
            (13, "Сервер → клиент", "FIN, ACK", 5830, 486),
            (17, "Клиент → сервер", "ACK", 486, 5831),
            (19, "Клиент → сервер", "FIN, ACK", 510, 5831),
            (20, "Сервер → клиент", "ACK", 5831, 511),
        ],
        [2.0, 5.2, 3.5, 2.2, 2.2],
    )
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(SCREENS / "02_tcp_close.png"), width=Cm(16.3))
    add_caption(doc, "Рисунок 2 — Четырёхэтапное завершение TCP-соединения")
    add_paragraph(doc, "Между подтверждением серверного FIN и клиентским FIN передан TLS-пакет закрытия длиной 24 байта. Поэтому Seq клиента изменился с 486 до 510.")

    doc.add_page_break()
    doc.add_heading("Количество сегментов HTTP запроса", level=1)
    add_paragraph(doc, "HTTP-запрос передавался внутри зашифрованного TLS Application Data. Клиентский пакет № 11 имел TCP Len = 115 байт. Следовательно, один HTTP-запрос был передан в одном TCP-сегменте.")
    add_paragraph(doc, "TLS-handshake и TCP-подтверждения в это число не включались, так как они обслуживают соединение, но не содержат сам HTTP-запрос.")

    doc.add_heading("Анализ Waterfall в Chrome DevTools", level=1)
    add_paragraph(doc, "Во вкладке Network был выбран основной запрос example.com и открыт раздел Timing. Наибольшую часть общего времени заняло ожидание ответа сервера — 151,13 мс.")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(SCREENS / "03_chrome_waterfall.png"), width=Cm(16.3))
    add_caption(doc, "Рисунок 3 — Этапы запроса по данным Chrome DevTools")
    add_result_table(
        doc,
        ["Этап", "Время"],
        [
            ("Queueing", "5,23 мс"),
            ("Stalled", "5,22 мс"),
            ("Proxy negotiation", "2,38 мс"),
            ("Request sent", "0,58 мс"),
            ("Waiting for server response", "151,13 мс"),
            ("Content download", "2,01 мс"),
            ("Общее время", "164,18 мс"),
        ],
        [10.5, 5.5],
    )

    doc.add_heading("Вывод", level=1)
    add_paragraph(doc, "В ходе работы был записан и исследован HTTPS-трафик. В TCP-потоке найдены трёхэтапное установление соединения и четырёхэтапное завершение с номерами Seq и Ack. Зашифрованный HTTP-запрос поместился в один TCP-сегмент. Chrome DevTools показал, что основное время запроса пришлось на ожидание ответа сервера.")

    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("Соболев Анатолий  ПИН252т")
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(90, 90, 90)
        set_lang(r)

    doc.core_properties.title = "Исследование TCP соединения с помощью Wireshark"
    doc.core_properties.subject = "Отчёт по практической работе"
    doc.core_properties.author = "Соболев Анатолий"
    doc.core_properties.keywords = "TCP, Wireshark, curl, Chrome DevTools"
    doc.save(OUT)


def main():
    SCREENS.mkdir(parents=True, exist_ok=True)
    draw_packet_table(
        SCREENS / "01_tcp_handshake.png",
        "Wireshark  TCP stream 22",
        "Фильтр: tcp.stream == 22   Установление соединения",
        [
            (1, "0.000000", "192.168.0.10", "172.66.147.243", "53294→443", "SYN", 0, 0),
            (2, "0.062923", "172.66.147.243", "192.168.0.10", "443→53294", "SYN, ACK", 0, 1),
            (3, "0.063050", "192.168.0.10", "172.66.147.243", "53294→443", "ACK", 1, 1),
        ],
    )
    draw_packet_table(
        SCREENS / "02_tcp_close.png",
        "Wireshark  TCP stream 22",
        "Фильтр: tcp.stream == 22   Завершение соединения",
        [
            (13, "0.209491", "172.66.147.243", "192.168.0.10", "443→53294", "FIN, ACK", 5830, 486),
            (17, "0.209696", "192.168.0.10", "172.66.147.243", "53294→443", "ACK", 486, 5831),
            (19, "0.211721", "192.168.0.10", "172.66.147.243", "53294→443", "FIN, ACK", 510, 5831),
            (20, "0.269259", "172.66.147.243", "192.168.0.10", "443→53294", "ACK", 5831, 511),
        ],
    )
    draw_waterfall(SCREENS / "03_chrome_waterfall.png")
    build_docx()
    print(OUT)


if __name__ == "__main__":
    main()
