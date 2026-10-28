"""Build the current teacher PDF directly from the first lecture HTML.

Requires ReportLab and an installed Arial or DejaVu Sans font family.
"""
from pathlib import Path
from html.parser import HTMLParser
import re,html
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,KeepTogether
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
root=Path(__file__).resolve().parents[1]
class SlideReader(HTMLParser):
 def __init__(self):
  super().__init__(); self.slides=[]; self.current=None; self.in_title=False; self.in_notes=False
 def handle_starttag(self,tag,attrs):
  attrs=dict(attrs)
  if tag=='section': self.current={'title':'','notes':''}
  if self.current is not None and tag in ('h1','h2'): self.in_title=True
  if self.current is not None and tag=='aside' and 'notes' in attrs.get('class','').split(): self.in_notes=True
 def handle_data(self,data):
  if self.current is not None:
   if self.in_title: self.current['title']+=data
   if self.in_notes: self.current['notes']+=data
 def handle_endtag(self,tag):
  if tag in ('h1','h2'): self.in_title=False
  if tag=='aside': self.in_notes=False
  if tag=='section' and self.current is not None:
   time,notes=self.current['notes'].split(' мин. ',1)
   self.current.update(time=time,notes=notes)
   self.slides.append(self.current);self.current=None
reader=SlideReader();reader.feed((root/'materials/Пара1/index.html').read_text())
slides=reader.slides
font_pairs=[
 (Path('/System/Library/Fonts/Supplemental/Arial.ttf'),Path('/System/Library/Fonts/Supplemental/Arial Bold.ttf')),
 (Path('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),Path('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))]
normal,bold=next((a,b) for a,b in font_pairs if a.exists() and b.exists())
pdfmetrics.registerFont(TTFont('Course',str(normal)))
pdfmetrics.registerFont(TTFont('CourseBold',str(bold)))
pdfmetrics.registerFontFamily('Course',normal='Course',bold='CourseBold',italic='Course',boldItalic='CourseBold')
navy=colors.HexColor('#102044');blue=colors.HexColor('#1168e8');pale=colors.HexColor('#eaf2ff')
styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name='CourseTitle',fontName='CourseBold',fontSize=23,leading=28,textColor=navy,spaceAfter=14))
styles.add(ParagraphStyle(name='CourseHeading',fontName='CourseBold',fontSize=13,leading=17,textColor=navy,spaceBefore=15,spaceAfter=7,keepWithNext=True))
styles.add(ParagraphStyle(name='CourseBody',fontName='Course',fontSize=10.5,leading=15,textColor=navy,spaceAfter=9))
styles.add(ParagraphStyle(name='CourseSmall',parent=styles['CourseBody'],fontSize=9,leading=12,textColor=colors.HexColor('#526986')))
styles.add(ParagraphStyle(name='CourseTime',parent=styles['CourseSmall'],textColor=blue,spaceAfter=5,keepWithNext=True))
P=lambda t,style='CourseBody':Paragraph(html.escape(t.replace('–','-').replace('—','-').replace('‑','-')),styles[style])
def tbl(data,widths):
 t=Table([[P(str(c),'CourseSmall') for c in row] for row in data],colWidths=widths,hAlign='LEFT')
 t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),pale),('LINEBELOW',(0,0),(-1,-1),.4,colors.HexColor('#c8d8eb')),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8)]))
 return t
story=[P('Первая пара: курс, оценивание и введение в ML','CourseTitle'),P('Сценарий преподавателя. 90 минут. 31 слайд','CourseTime'),P('Версия: 20 октября 2026. Порядок совпадает с основной HTML-презентацией. Восемь дополнительных слайдов находятся в appendix.html и не входят в обязательный маршрут.'),P('Маршрут занятия','CourseHeading'),tbl([
 ['Минуты', 'Слайды', 'Блок'],['0–10', '1–4', 'Знакомство и темы курса'],['10–17', '5–6', 'Зачем нужна база ML'],['17–33', '7–12', 'Баллы, задания и рейтинг'],['33–52', '13–19', 'Введение и типы задач'],['52–72', '20–25', 'Модели, переобучение и bias–variance'],['72–86', '26–30', 'Проверка, утечки и полный маршрут'],['86–90', '31', 'Итоги и вопросы']],[65,60,350]),P('Перед парой','CourseHeading'),P('Подготовьте представление на 60–90 секунд. Запустите презентацию через локальный HTTP-сервер, проверьте формулы, стрелки и режим докладчика S. Пример рейтинга на слайде 12 вымышленный; ссылку на рабочую таблицу выдайте отдельно после её создания.'),P('Критерии учебных медалей и частичного начисления 25/15 баллов публикуются до начала соревнований. Перед первым оцениваемым опросом проверьте Testportal на пробной попытке: 13 секунд на вопрос и блокировка при первом уходе со страницы.'),PageBreak()]
for i,s in enumerate(slides,1):
 story.append(P(f"{i}. {s['title']}",'CourseHeading'))
 story.append(P(f"{s['time']} мин",'CourseTime'))
 story.append(P(s['notes']))
story += [P('Дополнительные материалы','CourseHeading'),P('В appendix.html сохранены: математическая цель supervised learning, обозначения выборки, эмпирический риск, кластер и класс, предположения моделей, K-Means/DBSCAN/PCA, ловушки split и выбор метрики. Расширенные математические объяснения доступны в редактируемом КОНСПЕКТ_ЛЕКЦИИ.tex.'),P('Источники','CourseHeading'),P('Программа, система оценивания и правила опросов: документы syllabus/course_program.md, grading_and_homework.md и quiz_platform.md в репозитории курса.'),Paragraph('Официальные описания отбора: <link href="https://yandex.ru/jobs/interview/mldev?turbo=true" color="#1168e8">Яндекс</link> <link href="https://sberdevices.ru/career/ML/interview/" color="#1168e8">SberDevices</link>, <link href="https://www.tbank.ru/career/it/interview/ml/" color="#1168e8">Т-Банк</link>. Компетенции DS: <link href="https://github.com/avito-tech/playbook/blob/master/ds-skills.md" color="#1168e8">Авито</link>. Проверены 20 октября 2026. Учебные вопросы на слайде 6 не являются дословными вопросами компаний.',styles['CourseBody'])]
def page(c,d):
 c.saveState();c.setFont('Course',8);c.setFillColor(colors.HexColor('#526986'))
 c.drawString(42,A4[1]-27,'ДГТУ · Введение в машинное обучение · Первая пара')
 c.drawRightString(A4[0]-42,25,str(d.page));c.restoreState()
out=root/'output/pdf/КОНСПЕКТ_ЛЕКЦИИ.pdf'
out.parent.mkdir(parents=True,exist_ok=True)
doc=SimpleDocTemplate(str(out),pagesize=A4,rightMargin=52,leftMargin=52,topMargin=48,bottomMargin=43,title='Первая пара: сценарий преподавателя',author='Курс «Введение в машинное обучение»')
doc.build(story,onFirstPage=page,onLaterPages=page)
print(out)
