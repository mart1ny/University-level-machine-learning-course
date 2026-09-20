from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "docs" / "kratkiy_uchebnyy_plan.docx"

NAVY = "17365D"
PALE_BLUE = "EAF2F8"
PALE_GRAY = "F5F7F9"
BORDER = "D9D9D9"
BLACK = RGBColor(0, 0, 0)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=110, bottom=100, end=110):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_borders(cell, color=BORDER, size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_repeat_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Страница ")
    run.font.size = Pt(9)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])


def set_run_font(run, size=None, bold=None, color=None):
    run.font.name = "Arial"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Arial")
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Arial")
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def format_paragraph(paragraph, size=10.5, bold=False, align=None, after=2, line=1.05):
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    if align is not None:
        paragraph.alignment = align
    for run in paragraph.runs:
        set_run_font(run, size=size, bold=bold, color=BLACK)


def add_table(doc, headers, rows, widths, font_size=9.2, center_cols=()):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    header = table.rows[0]
    set_repeat_table_header(header)
    for idx, (cell, text, width) in enumerate(zip(header.cells, headers, widths)):
        cell.width = Inches(width)
        cell.text = text
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell, 110, 100, 110, 100)
        set_cell_borders(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, size=9.2, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=0, line=1.0)
        for run in p.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)

    for row_idx, values in enumerate(rows):
        row = table.add_row()
        prevent_row_split(row)
        fill = "FFFFFF" if row_idx % 2 == 0 else PALE_BLUE
        for col_idx, (cell, text, width) in enumerate(zip(row.cells, values, widths)):
            cell.width = Inches(width)
            cell.text = str(text)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_shading(cell, fill)
            set_cell_margins(cell)
            set_cell_borders(cell)
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.CENTER if col_idx in center_cols else WD_ALIGN_PARAGRAPH.LEFT
            format_paragraph(p, size=font_size, align=align, after=0, line=1.03)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(text, style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    format_paragraph(p, size=10.5, after=1, line=1.05)
    return p


def build_document():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.62)
    section.bottom_margin = Inches(0.62)
    section.left_margin = Inches(0.68)
    section.right_margin = Inches(0.68)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = BLACK

    title_style = styles["Title"]
    title_style.font.name = "Arial"
    title_style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    title_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    title_style.font.size = Pt(20)
    title_style.font.bold = True
    title_style.font.color.rgb = BLACK

    for name, size in (("Heading 1", 14), ("Heading 2", 11.5)):
        style = styles[name]
        style.font.name = "Arial"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = BLACK
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.keep_with_next = True

    header = section.header
    hp = header.paragraphs[0]
    hp.text = "ДГТУ  |  Введение в машинное обучение"
    format_paragraph(hp, size=8.5, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0, line=1.0)
    hp.runs[0].font.color.rgb = RGBColor(80, 80, 80)

    footer = section.footer
    fp = footer.paragraphs[0]
    set_repeat_page_number(fp)
    for run in fp.runs:
        set_run_font(run, size=8.5, color=RGBColor(90, 90, 90))

    p = doc.add_paragraph("Донской государственный технический университет")
    format_paragraph(p, size=10, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    p = doc.add_paragraph("Краткий учебный план дисциплины", style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    p = doc.add_paragraph("Введение в машинное обучение")
    format_paragraph(p, size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
    p = doc.add_paragraph("Рабочая версия")
    format_paragraph(p, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER, after=12)

    add_table(
        doc,
        ["Параметр", "Содержание"],
        [
            ("Уровень обучения", "3 курс бакалавриата"),
            ("Период", "Октябрь–декабрь; экзамен в январе"),
            ("Формат", "Онлайн"),
            ("Объём", "15 лекций и 15 практических занятий по 2 академических часа; всего 60 академических часов"),
            ("Итоговый контроль", "Экзамен"),
            ("Основные инструменты", "Python, NumPy, Pandas, Matplotlib, scikit-learn, CatBoost, Jupyter Notebook, GitHub"),
        ],
        [1.75, 5.25],
        font_size=9.6,
    )

    add_heading(doc, "Цель дисциплины", 1)
    p = doc.add_paragraph(
        "Сформировать у студентов фундаментальное понимание классического машинного обучения и базовые навыки решения задач на табличных данных — от постановки задачи и построения baseline до валидации, сравнения моделей и объяснения результатов."
    )
    format_paragraph(p, size=10.5, after=4, line=1.08)

    add_heading(doc, "Планируемые результаты обучения", 1)
    for item in (
        "понимать назначение основных алгоритмов классического машинного обучения и объяснять их работу;",
        "корректно разделять данные, выбирать метрики и проводить валидацию моделей;",
        "обучать и сравнивать модели с помощью scikit-learn;",
        "выявлять переобучение, дисбаланс классов и утечки данных;",
        "проводить воспроизводимый ML-эксперимент и формулировать выводы;",
        "отвечать на базовые вопросы по Classical ML на собеседовании уровня Junior Data Scientist или ML Intern.",
    ):
        add_bullet(doc, item)

    add_heading(doc, "Организация учебной недели", 1)
    add_table(
        doc,
        ["Занятие", "Назначение", "Типовая структура"],
        [
            ("Лекция", "Формирование понимания темы", "Задача → интуиция → визуализация → математика → алгоритм → код → эксперимент → вопрос собеседования"),
            ("Практика", "Закрепление и применение материала", "Разминка → совместный пример → самостоятельное решение → эксперимент → разбор ошибок → вывод"),
            ("Самостоятельная работа", "Проверка индивидуального понимания", "Домашние задания, Kaggle, итоговый проект и подготовка к квизам"),
        ],
        [1.05, 2.0, 3.95],
        font_size=9.2,
    )

    add_heading(doc, "Календарно тематический план", 1)
    p = doc.add_paragraph("Одна лекция и одна практическая пара в неделю.")
    format_paragraph(p, size=10, after=5)

    weeks = [
        (1, "Введение в ML", "Задачи ML, объект, признаки, target, supervised и unsupervised learning, общий ML pipeline", "Настройка Jupyter, загрузка данных, первый train/test split, baseline"),
        (2, "Данные и эксперимент", "Типы признаков, EDA, train/validation/test, loss и metric, основы data leakage", "EDA датасета, dummy baseline, таблица экспериментов, пример leakage"),
        (3, "KNN", "Расстояния, число соседей, scaling, curse of dimensionality, устройство fit и predict", "KNN на NumPy, выбор k, scaling, сравнение со sklearn"),
        (4, "Gradient Descent", "Loss, производная, градиент, learning rate, batch, SGD и mini-batch", "Реализация Gradient Descent, loss curve, эксперименты с learning rate"),
        (5, "Linear Regression", "Линейная модель, MSE, аналитическое и градиентное решения, MAE, RMSE, R²", "Linear Regression from scratch, метрики, выбросы, анализ коэффициентов"),
        (6, "Регуляризация", "Polynomial Features, сложность модели, Ridge, Lasso, ElasticNet, bias и variance", "Полиномиальная регрессия, Ridge/Lasso, исследование alpha и переобучения"),
        (7, "Logistic Regression", "Sigmoid, вероятность, Maximum Likelihood, LogLoss, decision boundary", "Sigmoid и LogLoss, Logistic Regression from scratch, predict_proba"),
        (8, "Метрики классификации", "Confusion matrix, precision, recall, F1, ROC-AUC, PR-AUC, threshold, imbalance", "Расчёт метрик, ROC/PR curves, выбор threshold, class weights"),
        (9, "Decision Tree", "Жадное построение, Gini, entropy, information gain, сложность дерева", "Поиск best split, визуализация дерева, влияние max_depth"),
        (10, "Validation и tuning", "K-Fold, Stratified K-Fold, learning curves, Grid Search и Randomized Search", "Cross-validation, Pipeline, GridSearchCV, начало Kaggle"),
        (11, "Bagging и Random Forest", "Bootstrap, bagging, случайные признаки, OOB, Extra Trees, feature importance", "Сравнение Tree, Bagging, Random Forest и Extra Trees"),
        (12, "Gradient Boosting", "Последовательное исправление ошибок, отрицательный градиент, XGBoost, LightGBM, CatBoost", "Упрощённый boosting, CatBoost, early stopping, продолжение Kaggle"),
        (13, "SVM и kernels", "Гиперплоскость, margin, support vectors, hinge loss, C, kernel trick, gamma", "Линейный и RBF SVM, эксперименты с C, gamma и scaling"),
        (14, "Признаки и Pipeline", "Feature engineering, ColumnTransformer, Pipeline, leakage, importance, основы SHAP", "Pipeline для смешанных данных, Leakage Challenge, ablation study"),
        (15, "Unsupervised Learning", "K-Means, DBSCAN, PCA, ограничения кластеризации, итоговая карта курса", "K-Means from scratch, сравнение с DBSCAN, PCA, консультация по проекту"),
    ]
    add_table(
        doc,
        ["Неделя", "Тема", "Содержание лекции", "Содержание практики"],
        weeks,
        [0.55, 1.35, 2.55, 2.55],
        font_size=8.6,
        center_cols=(0,),
    )

    add_heading(doc, "Система оценивания в разработке", 1)
    p = doc.add_paragraph(
        "Рейтинговая система и точное распределение баллов находятся в разработке. До утверждения документа приведённый ниже перечень отражает предполагаемые виды контроля без закрепления их удельного веса в итоговой оценке."
    )
    format_paragraph(p, size=10.5, after=5)
    grading = [
        ("Практические занятия", "Работа на занятиях, корректность решения и содержательный вывод", "В разработке"),
        ("Домашние задания", "Индивидуальные работы по основным разделам курса", "В разработке"),
        ("ML-квизы", "Короткие проверки понимания ключевых понятий", "В разработке"),
        ("Мини-собеседования", "Устные ответы по алгоритмам и прикладным ML-ситуациям", "В разработке"),
        ("Kaggle", "Валидация, модели, эксперименты и итоговый отчёт", "В разработке"),
        ("Итоговый проект", "Полный ML pipeline и краткая защита", "В разработке"),
        ("Экзамен", "Теория, прикладная ML-ситуация и анализ результатов", "В разработке"),
    ]
    add_table(doc, ["Вид работы", "Что предполагается оценивать", "Статус"], grading, [1.75, 4.15, 1.1], font_size=9.2, center_cols=(2,))

    add_heading(doc, "Домашние задания", 1)
    homework = [
        ("HW1", "Первый корректный ML-эксперимент", "Уточняются"),
        ("HW2", "KNN from scratch и scaling", "Уточняются"),
        ("HW3", "Linear Regression, Gradient Descent и регуляризация", "Уточняются"),
        ("HW4", "Logistic Regression, метрики и threshold", "Уточняются"),
        ("HW5", "Decision Tree, Cross-Validation и Random Forest", "Уточняются"),
        ("HW6", "Feature Engineering, Pipeline и Leakage Challenge", "Уточняются"),
        ("HW7", "Clustering и PCA", "Уточняются"),
    ]
    add_table(doc, ["Работа", "Предварительная тема", "Баллы"], homework, [0.8, 4.95, 1.25], font_size=9.4, center_cols=(0, 2))

    add_heading(doc, "Основные правила оценивания", 1)
    for item in (
        "Практика оценивается за выполненную работу, а не только за присутствие.",
        "Домашние задания выполняются индивидуально и сдаются через персональные GitHub-репозитории.",
        "Студент должен уметь объяснить любой фрагмент сданного решения.",
        "Использование AI-инструментов допускается при раскрытии способа использования и самостоятельной проверке результата.",
        "Точное количество работ, сроки, максимальные баллы и правила пересдачи будут утверждены отдельно.",
    ):
        add_bullet(doc, item)

    p = doc.add_paragraph(
        "Методический ориентир курса — «Тренировки по ML 1.0» Яндекса. Материалы курса адаптируются к университетскому формату и уровню подготовки студентов; структура и задания не копируются дословно."
    )
    p.paragraph_format.space_before = Pt(6)
    format_paragraph(p, size=9.5, after=0, line=1.05)

    doc.core_properties.title = "Краткий учебный план дисциплины Введение в машинное обучение"
    doc.core_properties.subject = "Учебный план дисциплины ДГТУ"
    doc.core_properties.author = "ДГТУ"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
