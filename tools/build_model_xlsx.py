# -*- coding: utf-8 -*-
"""
Допоміжна фінансова модель (окремий файл, не подається): припущення, пакети послуг, помісячний план,
покриття гранту податками, сценарії, оцінка балів за додатком 3 до наказу Мінекономіки № 2.
Запуск: python3 tools/build_model_xlsx.py <вихідний_xlsx>
"""
import sys, os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
sys.path.insert(0, os.path.dirname(__file__))
import model as M

OUT = sys.argv[1]
wb = openpyxl.Workbook()
H = Font(bold=True); HF = PatternFill("solid", fgColor="DDEBF7"); INP = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BBBBBB"); BR = Border(left=thin, right=thin, top=thin, bottom=thin)
def hdr(ws, row, values, widths=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="center"); c.border = BR
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w

# 1. Припущення
ws = wb.active; ws.title = "Припущення"
ws["A1"] = "Фінансова модель бізнес-плану ФОП Федоренко Т. А., версія 2 (допоміжний файл — не подається до заяви)"; ws["A1"].font = Font(bold=True, size=12)
ws["A2"] = "Жовті клітинки — вхідні дані, які можна змінювати; решта аркушів перераховуються автоматично."
rows = [
    ("Сума гранту, грн", M.GRANT, "п. 6 Порядку: 500 000–1 500 000 для масштабування без надбавок"),
    ("Строк проєкту, місяців", M.TERM, "24 / 36 / 60"),
    ("Ставка ЄП (3 група, без ПДВ)", M.EP_RATE, "п. 293.3 ПКУ"),
    ("Військовий збір ФОП 3 групи", M.VZ_FOP, "1 % доходу (п. 16-1 підрозд. 10 розд. XX ПКУ)"),
    ("ПДФО із зарплати", M.PDFO, "18 %"),
    ("Військовий збір із зарплати", M.VZ_EMP, "5 %"),
    ("ЄСВ роботодавця", M.ESV, "22 %"),
    ("Мінімальна зарплата (прогноз 2027)", M.MIN_WAGE_2027, "проєкт Держбюджету-2027; у 2026 році — 8 647 грн"),
    ("ЄСВ ФОП за себе, грн/міс", "=B10*B9", "22 % мінімальної зарплати"),
    ("Зарплата: консультант з комплаєнсу (рік 1 / 2 / 3)", "25 000 / 27 500 / 30 000", "нарахована, грн/міс"),
    ("Зарплата: аналітик-адміністратор (рік 1 / 2 / 3)", "20 000 / 22 500 / 25 000", "нарахована, грн/міс"),
    ("Оренда офісу, грн/міс", 15_000, "М1–М12 сплачено наперед з гранту; далі з доходу"),
    ("Власні кошти: навчання ICA, грн", 385_860, "110 000 у М1 + 275 860 у М13"),
    ("Власні кошти: оборотний резерв, грн", 100_000, "М1; підтвердити випискою"),
]
hdr(ws, 3, ["Параметр", "Значення", "Коментар"], [52, 18, 70])
for i, (a, b, c) in enumerate(rows, 4):
    ws.cell(row=i, column=1, value=a); cc = ws.cell(row=i, column=2, value=b); cc.fill = INP; ws.cell(row=i, column=3, value=c)
    if isinstance(b, float) and b < 1: cc.number_format = "0.0%"
    elif isinstance(b, (int, float)): cc.number_format = "#,##0.00" if isinstance(b, float) else "#,##0"
ws["B12"].number_format = "#,##0.00"

# 2. Пакети послуг і план продажів
ws2 = wb.create_sheet("Пакети і продажі")
hdr(ws2, 1, ["Код", "Пакет послуг", "Ціна, грн", "Норматив годин", "Результат для клієнта"], [6, 70, 12, 14, 60])
hours = {"A": 30, "B": 50, "C": 2.5, "D": 12, "E": 19}
results = {"A": "звіт про відповідність процесів з переліком ризиків і планом усунення",
           "B": "затверджена антикорупційна програма / політики, регламенти, наказ про уповноваженого",
           "C": "звіт з реєстрів ЄДР, судових рішень, санкційних списків, податкового боргу",
           "D": "проведений тренінг, матеріали, тест, протокол навчання",
           "E": "щомісячний моніторинг, перевірки контрагентів, консультації, підготовка до перевірок"}
for i, k in enumerate("ABCDE", 2):
    ws2.cell(row=i, column=1, value=k); ws2.cell(row=i, column=2, value=M.PACKAGES[k][0])
    c = ws2.cell(row=i, column=3, value=M.PACKAGES[k][1]); c.number_format = "#,##0"; c.fill = INP
    ws2.cell(row=i, column=4, value=hours[k]); ws2.cell(row=i, column=5, value=results[k])
r0 = 9
hdr(ws2, r0, ["Місяць", "A, шт", "B, шт", "C, шт", "D, шт", "E, клієнтів", "Дохід основні (A+B+E)", "Дохід додаткові (C+D)", "Разом дохід, грн", "Годин навантаження"])
for m in range(1, M.TERM + 1):
    k = M.counts_for_month(m); r = r0 + m
    ws2.cell(row=r, column=1, value=m)
    for j, key in enumerate("ABCDE", 2):
        c = ws2.cell(row=r, column=j, value=k[key]); c.fill = INP
    ws2.cell(row=r, column=7, value=f"=B{r}*$C$2+C{r}*$C$3+F{r}*$C$6").number_format = "#,##0"
    ws2.cell(row=r, column=8, value=f"=D{r}*$C$4+E{r}*$C$5").number_format = "#,##0"
    ws2.cell(row=r, column=9, value=f"=G{r}+H{r}").number_format = "#,##0"
    ws2.cell(row=r, column=10, value=f"=B{r}*$D$2+C{r}*$D$3+D{r}*$D$4+E{r}*$D$5+F{r}*$D$6").number_format = "#,##0.0"
rt = r0 + M.TERM + 1
ws2.cell(row=rt, column=1, value="Разом").font = H
for col in (7, 8, 9):
    ws2.cell(row=rt, column=col, value=f"=SUM({get_column_letter(col)}{r0+1}:{get_column_letter(col)}{rt-1})").number_format = "#,##0"
ws2.cell(row=rt + 1, column=1, value="Довідково: 3 фахівці × ~120 оплачуваних годин/міс = ~360 год; максимальне навантаження у моделі перевірте у колонці J.")

# 3. Помісячний план
ws3 = wb.create_sheet("Помісячний план")
cols = ["Місяць", "Дохід від послуг", "Грант", "Власні кошти", "Разом надходжень", "Грантові закупівлі", "Навчання (власні)", "Оренда", "Інші поточні витрати",
        "Зарплата нарахована", "ЄСВ роботодавця", "ПДФО", "ВЗ працівників", "ЄП 5 %", "ВЗ ФОП 1 %", "ЄСВ ФОП", "Вилучення власниці", "Разом виплат",
        "Залишок на кінець місяця", "Податки та ЄСВ за місяць", "Накопичені податки", "Залишок зобов'язання за грантом"]
hdr(ws3, 1, cols, [8] + [16] * (len(cols) - 1))
months = M.build_months()
training = {mm: v for (n, v, mm) in M.OWN if "Навчання" in n}
for x in months:
    r = x.m + 1
    capex = M.GRANT if x.m == 1 else 0
    other = x.opex_total - capex - training.get(x.m, 0) - x.expenses.get(34, 0)
    vals = [x.m, x.revenue, x.grant_in, x.own_in, None, capex, training.get(x.m, 0), x.expenses.get(34, 0), other, x.payroll_new + x.payroll_exist,
            None, None, None, None, None, M.ESV_FOP_MONTH, x.drawings]
    for j, v in enumerate(vals, 1):
        c = ws3.cell(row=r, column=j, value=v)
        if j in (2, 3, 4, 6, 7, 8, 9, 10, 16, 17): c.fill = INP
    ws3.cell(row=r, column=5, value=f"=B{r}+C{r}+D{r}")
    ws3.cell(row=r, column=11, value=f"=J{r}*Припущення!$B$10")
    ws3.cell(row=r, column=12, value=f"=J{r}*Припущення!$B$8")
    ws3.cell(row=r, column=13, value=f"=J{r}*Припущення!$B$9")
    ws3.cell(row=r, column=14, value=f"=B{r}*Припущення!$B$6")
    ws3.cell(row=r, column=15, value=f"=B{r}*Припущення!$B$7")
    ws3.cell(row=r, column=18, value=f"=F{r}+G{r}+H{r}+I{r}+J{r}+K{r}+N{r}+O{r}+P{r}+Q{r}")
    ws3.cell(row=r, column=19, value=f"=E{r}-R{r}" if x.m == 1 else f"=S{r-1}+E{r}-R{r}")
    ws3.cell(row=r, column=20, value=f"=K{r}+L{r}+M{r}+N{r}+O{r}+P{r}")
    ws3.cell(row=r, column=21, value=f"=T{r}" if x.m == 1 else f"=U{r-1}+T{r}")
    ws3.cell(row=r, column=22, value=f"=MAX(0,Припущення!$B$4-U{r})")
    for j in range(2, 23): ws3.cell(row=r, column=j).number_format = "#,##0"
rt = M.TERM + 2
ws3.cell(row=rt, column=1, value="Разом").font = H
for j in list(range(2, 19)) + [20]:
    L = get_column_letter(j); ws3.cell(row=rt, column=j, value=f"=SUM({L}2:{L}{rt-1})").number_format = "#,##0"
ws3.cell(row=rt + 2, column=1, value="Покриття гранту податками (місяць, коли накопичені податки ≥ грант):")
ws3.cell(row=rt + 2, column=5, value=f'=IFERROR(MATCH(0,V2:V{rt-1},0),"не покрито")')
ws3.cell(row=rt + 3, column=1, value="Мінімальний залишок коштів за 36 місяців, грн:")
ws3.cell(row=rt + 3, column=5, value=f"=MIN(S2:S{rt-1})").number_format = "#,##0"
ws3.freeze_panes = "B2"

# 4. Сценарії
ws4 = wb.create_sheet("Сценарії")
hdr(ws4, 1, ["Сценарій", "Дохід за 36 міс", "Податки та ЄСВ за 36 міс", "Покриття гранту", "Мінімальний залишок коштів", "Операційний прибуток", "Коментар"], [38, 18, 22, 16, 22, 20, 60])
def scen(rev_mult=1.0, pay_mult=1.0, delay=0):
    tot_rev = 0; taxes = 0; bal = 0; min_bal = None; cover = None; cum = 0; op = 0
    for x in months:
        rev = x.revenue * rev_mult if x.m > delay else 0
        pay = (x.payroll_new + x.payroll_exist) * pay_mult
        t = pay * (M.PDFO + M.VZ_EMP + M.ESV) + rev * (M.EP_RATE + M.VZ_FOP) + M.ESV_FOP_MONTH
        out = x.opex_total + pay + x.drawings + pay * M.ESV + rev * (M.EP_RATE + M.VZ_FOP) + M.ESV_FOP_MONTH
        bal += rev + x.grant_in + x.own_in - out
        min_bal = bal if min_bal is None else min(min_bal, bal)
        cum += t; taxes += t; tot_rev += rev
        if cover is None and cum >= M.GRANT: cover = x.m
        capex = M.GRANT if x.m == 1 else 0
        op += rev - (x.opex_total - capex - training.get(x.m, 0) + pay + pay * M.ESV + rev * (M.EP_RATE + M.VZ_FOP) + M.ESV_FOP_MONTH)
    return tot_rev, taxes, cover, min_bal, op
scenarios = [
    ("Базовий", scen(), "консервативний план продажів; 2 працівники з М1"),
    ("Дохід −20 %", scen(rev_mult=0.8), "перевірка умови «податки ≥ грант» за падіння виручки"),
    ("Дохід −30 %", scen(rev_mult=0.7), "податки все ще ≥ гранту, але касовий розрив у перші місяці — потрібен більший резерв або менші вилучення"),
    ("Старт продажів на 3 місяці пізніше", scen(delay=3), "витрати на команду з М1, дохід з М4; показує потребу в резерві"),
    ("Зарплати +15 %", scen(pay_mult=1.15), "податки зростають, умова виконується з запасом; тисне на прибуток"),
]
for i, (name, (rev, tx, cov, mb, op), comm) in enumerate(scenarios, 2):
    ws4.cell(row=i, column=1, value=name)
    ws4.cell(row=i, column=2, value=round(rev)).number_format = "#,##0"
    ws4.cell(row=i, column=3, value=round(tx)).number_format = "#,##0"
    ws4.cell(row=i, column=4, value=(f"М{cov}" if cov else "не покрито за 36 міс"))
    ws4.cell(row=i, column=5, value=round(mb)).number_format = "#,##0"
    ws4.cell(row=i, column=6, value=round(op)).number_format = "#,##0"
    ws4.cell(row=i, column=7, value=comm)
ws4.cell(row=9, column=1, value="Висновок: у всіх сценаріях сума сплачених податків і ЄСВ за 36 місяців перевищує грант 1 000 000 грн (запас у базовому — 28 %). "
                               "Слабке місце — ліквідність у М2–М6 за падіння виручки або пізнього старту: тримати резерв не менше 150 000 грн і мінімальні вилучення в перший рік.")

# 5. Оцінка балів (додаток 3 до наказу № 2)
ws5 = wb.create_sheet("Оцінка балів (дод. 3)")
hdr(ws5, 1, ["№", "Критерій (для фізичних осіб / ФОП)", "Варіанти та бали з коефіцієнтом", "Що впливає / що може зробити заявниця", "Обраний варіант (бали)"], [5, 44, 60, 60, 20])
crit = [
    (1, "Кількість погашених кредитів", "0 — 0; 1–3 — 10; 4–5 — 14; >5 — 20", "Факт кредитної історії. Не брати кредити заради балів.", 0),
    (2, "Максимальний строк прострочення (сума > 1000 грн, за 3 роки)", "0 днів — 20; 1–15 — 14; 16–30 — 10; 31–60 — 2; >60 — 0; позик не було — 10", "Погасити будь-які прострочення до подання заяви; взяти кредитний звіт в УБКІ/ПВБКІ.", 10),
    (3, "Кількість прострочень < 90 днів за 2 роки", "0 — 20; 1–3 — 10; >3 — 0", "", 20),
    (4, "Кількість запитів на кредит за 12 міс", "<2 — 15; 3 — 11; 4 — 5; ≥5 — 0", "Не подавати кредитних заявок за 12 місяців до подання.", 15),
    (5, "Місячний платіж за кредитами / дохід", "<0,1 — 14; 0,1–0,24 — 10; 0,25–0,5 — 7; >0,5 — 0", "", 14),
    (6, "Створення нових робочих місць", "0 — 0; 1 — 1; 2 — 3; >2 — 4", "У плані 2 нових місця — 3 бали.", 3),
    (7, "Очікувана рентабельність за бізнес-планом", ">0,5 — 6; 0,3–0,5 — 4; 0,15–0,3 — 3; <0,15 — 0", "У плані 67 % — 6 балів.", 6),
    (8, "Період окупності, років", "<1 — 4; ≈3 — 3; ≥5 — 0", "У плані 1,7 року — очікувано 3–4 бали.", 3),
    (9, "Участь власними коштами", ">0,5 — 20; 0,3–0,5 — 4 (у наказі, ймовірно, 14); 0,15–0,3 — 3; <0,15 — 0", "У плані 33 % — підтвердити випискою; збільшення власної участі до >50 % дало б 20 балів.", 4),
    (10, "Нерухомість для бізнесу", "не потрібна — 10; власна (у т. ч. подружжя) — 10; родичів — 7; оренда >1 року — 4; оренда <1 року — 0", "Укласти договір оренди на 12+ місяців до подання (4 бали) або використати власне/родинне приміщення (10/7).", 4),
    (11, "Транспорт для бізнесу", "не потрібен — 10; власний — 7; родичів — 5; орендований — 0", "Для консалтингу транспорт не потрібен — 10.", 10),
    (12, "Сімейний стан", "одружена — 4; розлучена — 2; не у шлюбі / цивільний шлюб — 0", "Факт.", None),
    (13, "Утриманці (неповнолітні діти)", "немає — 10; 1 — 5; 2+ — 0", "Факт.", None),
    (14, "Активи у власності", "житло у власності (у т. ч. подружжя) — 10; у родичів — 7; оренда — 0; авто у власності — 5; авто родичів — 3", "Факт; зазначити у заяві все, що є.", None),
    (15, "Вік", "18–25 — 2; 26–45 — 3; 46–55 — 4; >55 — 2", "Факт.", None),
    (16, "Стаж за попереднім/поточним місцем роботи", "<5 років — 0; 5–12 — 2; >12 — 4", "Стаж директора ліцею — зазначити у заяві.", None),
    (17, "Користування послугами банку", "немає рахунків — 0; поточний/картковий — 2; кредити — 3; депозит — 5", "Відкрити депозит в обраному уповноваженому банку до подання — 5 балів.", 5),
    (18, "Освіта", "бакалавр — 2; магістр/спеціаліст — 3; аспірантура/PhD — 4; профтехосвіта — 2; 11 класів — 1", "Факт; вказати найвищий рівень.", None),
    (19, "Тип населеного пункту", "<20 тис. — 0; 20–200 тис. — 1; 200–500 тис. — 2; 0,5–1 млн — 3; >1 млн — 3; >2 млн — 4", "Жашків (<20 тис.) — 0. Не змінювати місце реєстрації заради балів.", 0),
    (20, "Цивільні судові справи", "є — 0; немає — 2", "Перевірити себе в реєстрі судових рішень та YouControl.", 2),
    (21, "Ризик виду діяльності (експертне судження банку)", "низький — 10; середній — 5; високий — 0", "Консалтинг — зазвичай низький/середній.", 5),
]
for i, (n, c, opts, lever, chosen) in enumerate(crit, 2):
    ws5.cell(row=i, column=1, value=n); ws5.cell(row=i, column=2, value=c); ws5.cell(row=i, column=3, value=opts); ws5.cell(row=i, column=4, value=lever)
    cc = ws5.cell(row=i, column=5, value=chosen); cc.fill = INP
    for j in range(1, 6): ws5.cell(row=i, column=j).alignment = Alignment(wrap_text=True, vertical="top")
last = len(crit) + 1
ws5.cell(row=last + 1, column=2, value="Разом балів банку (максимум 200; прохідний поріг 99)").font = H
ws5.cell(row=last + 1, column=5, value=f"=SUM(E2:E{last})").font = H
ws5.cell(row=last + 2, column=2, value="Співбесіда (0–50): мотивація, розуміння бізнес-плану, маркетингового плану, операційного плану, релевантний досвід — по 0/3/5/7/10").font = H
ws5.cell(row=last + 2, column=5, value=35).fill = INP
ws5.cell(row=last + 3, column=2, value="Підсумковий бал (поріг 130 з 250)").font = H
ws5.cell(row=last + 3, column=5, value=f"=E{last+1}+E{last+2}").font = H
ws5.cell(row=last + 5, column=2, value="Порожні клітинки у колонці E — особисті дані, які відомі лише заявниці; заповніть їх, щоб побачити власну оцінку. Джерело: додаток 3 і додаток 5 до наказу Мінекономіки від 22.07.2026 № 2 (зареєстровано Мін'юстом 28.08.2026 за № 1185/46579).")

# 6. Ліміти п. 7
ws6 = wb.create_sheet("Ліміти п.7")
hdr(ws6, 1, ["Стаття інвестиційного плану", "Сума, грн", "Частка гранту", "Ліміт п. 7 Порядку", "Статус"], [60, 14, 14, 40, 12])
lim = {"Ліцензійне": 0.25, "Вебсайт": 0.10, "Налаштування": 0.20, "Маркетинг": 0.10, "Фірмовий": 0.50, "Оренда": 0.25}
for i, (name, amt, just, l) in enumerate(M.INVEST, 2):
    ws6.cell(row=i, column=1, value=name); ws6.cell(row=i, column=2, value=amt).number_format = "#,##0"
    ws6.cell(row=i, column=3, value=f"=B{i}/Припущення!$B$4").number_format = "0.0%"
    key = next((k for k in lim if name.startswith(k)), None)
    ws6.cell(row=i, column=4, value=(f"не більше {lim[key]:.0%}" if key else "без окремого ліміту"))
    ws6.cell(row=i, column=5, value=(f'=IF(C{i}<={lim[key]},"у межах","ПЕРЕВИЩЕНО")' if key else "у межах"))
r = len(M.INVEST) + 2
ws6.cell(row=r, column=1, value="Разом").font = H; ws6.cell(row=r, column=2, value=f"=SUM(B2:B{r-1})").number_format = "#,##0"
ws6.cell(row=r, column=5, value=f'=IF(B{r}=Припущення!$B$4,"= грант","≠ грант")')
wb.save(OUT); print("saved", OUT)
