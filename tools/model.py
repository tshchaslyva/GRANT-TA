# -*- coding: utf-8 -*-
"""
Фінансова модель відкоригованого бізнес-плану ФОП Федоренко Т. А. (версія 2).
Єдине джерело цифр для xlsx-форми, звіту та перевірки.
Усі суми у гривнях. М1 = місяць зарахування гранту (орієнтовно травень 2027).
"""
from dataclasses import dataclass, field

TERM = 36
GRANT = 1_000_000
M1_LABEL = "травень 2027"          # орієнтовний місяць зарахування гранту
M1_YEAR, M1_MONTH = 2027, 5

# --- Податкові ставки (чинні на 10.2026; ЄСВ ФОП — з проєкту бюджету-2027: 9 546 грн × 22 %) ---
PDFO = 0.18
VZ_EMP = 0.05
ESV = 0.22
EP_RATE = 0.05        # ЄП 3 група, неплатник ПДВ
VZ_FOP = 0.01         # військовий збір ФОП 3 групи (1 % доходу)
MIN_WAGE_2027 = 9_546
ESV_FOP_MONTH = round(MIN_WAGE_2027 * ESV, 2)   # 2 100,12

# --- Пакети послуг (ціни без ПДВ; ФОП — неплатник ПДВ) ---
PACKAGES = {
    "A": ("Комплаєнс-діагностика (експрес-оцінка процесів, звіт з рекомендаціями)", 25_000),
    "B": ("Антикорупційна програма / пакет політик і процедур «під ключ»", 45_000),
    "C": ("Перевірка контрагента (звіт з відкритих реєстрів і санкційних списків)", 2_500),
    "D": ("Навчання персоналу клієнта (тренінг / вебінар, 4 год)", 12_000),
    "E": ("Абонентський супровід (комплаєнс-офіцер на аутсорсингу), на місяць", 15_000),
}

def counts_for_month(m: int):
    """Кількість проданих послуг у місяці m (1..36). Консервативний базовий сценарій."""
    # абонентські клієнти (E)
    if m <= 2: e = 1
    elif m <= 5: e = 2
    elif m <= 8: e = 3
    elif m <= 18: e = 4
    elif m <= 24: e = 5
    else: e = 6
    # навчання (D)
    d = 1 if m <= 24 else 2
    # перевірки контрагентів (C), звітів на місяць
    if m <= 6: c = 4
    elif m <= 12: c = 6
    elif m <= 24: c = 8
    else: c = 10
    # проєкти: діагностика (A) та політики (B)
    if m <= 5: a, b = 1, 0
    elif m <= 12: a, b = 1, 1
    else: a, b = 2, 1
    return dict(A=a, B=b, C=c, D=d, E=e)

def revenue_split(m: int):
    k = counts_for_month(m)
    p = {key: PACKAGES[key][1] for key in PACKAGES}
    main = k["A"]*p["A"] + k["B"]*p["B"] + k["E"]*p["E"]      # рядок 16: основні послуги
    extra = k["C"]*p["C"] + k["D"]*p["D"]                      # рядок 17: додаткові послуги
    return main, extra

# --- Кадровий план: 2 нових робочих місця з М1 (нарахована зарплата, грн/міс) ---
def payroll_new(m: int):
    if m <= 12: return 25_000 + 20_000
    if m <= 24: return 27_500 + 22_500
    return 30_000 + 25_000
PAYROLL_EXISTING = 0   # працівників на дату подання заяви немає

# --- Інвестиційний план за рахунок гранту (п. 7 Порядку) ---
INVEST = [
    # (стаття, сума, обґрунтування, напрям п.7 / ліміт)
    ("Комп'ютерна техніка та обладнання робочих місць", 300_000,
     "ноутбук керівника 1×55 000; ноутбуки працівників 2×42 000; монітори 27\" 3×9 000; БФП лазерний А4 22 000; "
     "зарядна станція/ДБЖ 1 кВт·год 35 000; мережеве обладнання і NAS для резервних копій 25 000; гарнітури/вебкамери 12 000; "
     "екран 55\" для презентацій і навчань 20 000; шредер 6 000; сейф для документів 14 000",
     "пп. 1 та 5 п. 7 (обладнання, облаштування робочих місць) — без окремого ліміту"),
    ("Меблі для 3 робочих місць і переговорної зони", 90_000,
     "столи 3×7 000; крісла 3×6 000; шафи для документів 2×8 000; стіл переговорний з 6 стільцями 25 000; тумби та полиці 10 000",
     "пп. 1 та 5 п. 7 — без окремого ліміту"),
    ("Ліцензійне програмне забезпечення (12 місяців)", 180_000,
     "LIGA360 (2 користувачі) 72 000; YouControl (перевірка контрагентів) 48 000; Microsoft 365 Business (3 користувачі) 15 000; "
     "CRM-система 18 000; електронний документообіг і КЕП 6 000; антивірус і резервне копіювання 9 000; сервіс відеоконференцій 12 000",
     "пп. 3 п. 7 — не більше 25 % гранту (250 000)"),
    ("Вебсайт із базою знань та інтеграцією з CRM", 95_000,
     "розроблення сайту: каталог послуг, форма замовлення, база знань, інтеграція з CRM, захист даних",
     "пп. 3 п. 7 — не більше 10 % гранту (100 000)"),
    ("Налаштування та інтеграція цифрових рішень", 40_000,
     "налаштування CRM, електронного документообігу, резервного копіювання, політик кібербезпеки робочих місць",
     "пп. 3 п. 7 — не більше 20 % гранту (200 000)"),
    ("Маркетинг та реклама", 95_000,
     "таргетована реклама у LinkedIn/Meta/Google 60 000; публікації у профільних медіа та участь у бізнес-подіях 35 000",
     "пп. 4 п. 7 — не більше 10 % гранту (100 000)"),
    ("Фірмовий стиль і промоматеріали", 20_000,
     "айдентика, шаблони презентацій і комерційних пропозицій, друк матеріалів",
     "пп. 4 п. 7 — не більше 50 % гранту"),
    ("Оренда офісного приміщення (12 місяців)", 180_000,
     "м. Жашків, вул. Соборна, 37: 12 місяців × 15 000 грн за договором оренди строком не менше 12 місяців",
     "пп. 2 п. 7 — не більше 25 % гранту (250 000)"),
]
assert sum(x[1] for x in INVEST) == GRANT, sum(x[1] for x in INVEST)

# --- Власні кошти ---
OWN = [
    ("Навчання: Розширений сертифікат ICA з управління, ризиків та комплаєнсу", 110_000, 1),
    ("Навчання: Диплом ICA з управління, ризиків та комплаєнсу", 275_860, 13),
    ("Оборотні кошти (резерв на перші місяці)", 100_000, 1),
]
OWN_TOTAL = sum(x[1] for x in OWN)

# --- Поточні витрати (грн/міс) ---
def opex(m: int):
    """Повертає словник: рядок шаблону -> сума за місяць (без грантових закупівель М1)."""
    season_cold = ((M1_MONTH - 1 + (m - 1)) % 12 + 1) in (11, 12, 1, 2, 3)
    d = {}
    d[33] = 1_500                                   # канцтовари, витратні матеріали
    d[34] = 0 if m <= 12 else 15_000                # оренда (М1–М12 сплачено наперед з гранту)
    d[35] = 0
    d[36] = 300 + sum(x[1] for x in OWN if x[2] == m and "Навчання" in x[0])   # банк + навчання за власні кошти
    d[37] = 4_000 if season_cold else 2_000         # комунальні
    d[38] = 0
    d[39] = 0 if m == 1 else (3_000 if m <= 12 else 5_000)   # маркетинг після стартового бюджету
    d[40] = 1_200                                   # зв'язок, інтернет
    d[41] = 0
    d[42] = 0
    d[43] = 2_000                                   # відрядження/транспорт
    d[44] = 3_000 + (0 if m <= 12 else 15_000)      # бухгалтерський супровід + продовження ліцензій з М13
    return d

def grant_items_by_row():
    """Грантові закупівлі М1 за рядками шаблону."""
    return {
        34: 180_000,                 # оренда
        39: 95_000 + 20_000,         # маркетинг + фірмовий стиль
        41: 300_000 + 90_000,        # обладнання + меблі
        44: 180_000 + 95_000 + 40_000,   # ПЗ + сайт + налаштування
    }

def fop_drawings(m: int):
    if m <= 3: return 0
    if m <= 12: return 10_000
    if m <= 24: return 25_000
    return 35_000

@dataclass
class Month:
    m: int
    rev_main: float
    rev_extra: float
    grant_in: float
    own_in: float
    expenses: dict = field(default_factory=dict)   # рядок -> сума
    payroll_new: float = 0
    payroll_exist: float = 0
    drawings: float = 0
    @property
    def revenue(self): return self.rev_main + self.rev_extra
    @property
    def inflow(self): return self.revenue + self.grant_in + self.own_in
    @property
    def pdfo(self): return round((self.payroll_new + self.payroll_exist) * PDFO, 2)
    @property
    def vz_emp(self): return round((self.payroll_new + self.payroll_exist) * VZ_EMP, 2)
    @property
    def esv_emp(self): return round((self.payroll_new + self.payroll_exist) * ESV, 2)
    @property
    def ep(self): return round(self.revenue * EP_RATE, 2)
    @property
    def vz_fop(self): return round(self.revenue * VZ_FOP, 2)
    @property
    def esv_fop(self): return ESV_FOP_MONTH
    @property
    def taxes_total(self):
        return round(self.pdfo + self.vz_emp + self.esv_emp + self.ep + self.vz_fop + self.esv_fop, 2)
    @property
    def opex_total(self): return sum(self.expenses.values())
    @property
    def payroll_total(self): return self.payroll_new + self.payroll_exist + self.drawings
    @property
    def outflow(self):
        # як у шаблоні: рядок 72 = витрати(32) + оплата праці(46) + ЄСВ(54,58) + ЄП, ВЗ ФОП, ЄСВ ФОП (60,61,63-65,67-69)
        return round(self.opex_total + self.payroll_total + self.esv_emp + self.ep + self.vz_fop + self.esv_fop, 2)

def build_months():
    months = []
    for m in range(1, TERM + 1):
        main, extra = revenue_split(m)
        exp = opex(m)
        if m == 1:
            for r, v in grant_items_by_row().items():
                exp[r] = exp.get(r, 0) + v
        own_in = sum(x[1] for x in OWN if x[2] == m)
        months.append(Month(m, main, extra, GRANT if m == 1 else 0, own_in, exp,
                            payroll_new(m), PAYROLL_EXISTING, fop_drawings(m)))
    return months

def summary(months=None):
    months = months or build_months()
    rev = sum(x.revenue for x in months)
    taxes = sum(x.taxes_total for x in months)
    taxes_from_m2 = sum(x.taxes_total for x in months if x.m >= 2)
    cum = 0; cover_month = None; cover_month_m2 = None; cum2 = 0
    for x in months:
        cum += x.taxes_total
        if cover_month is None and cum >= GRANT: cover_month = x.m
        if x.m >= 2:
            cum2 += x.taxes_total
            if cover_month_m2 is None and cum2 >= GRANT: cover_month_m2 = x.m
    inflow = sum(x.inflow for x in months); outflow = sum(x.outflow for x in months)
    balance = 0; min_bal = None; min_bal_m = None; balances = []
    for x in months:
        balance += x.inflow - x.outflow
        balances.append(balance)
        if min_bal is None or balance < min_bal: min_bal, min_bal_m = balance, x.m
    payroll_new_total = sum(x.payroll_new for x in months)
    # операційний результат (без грантових закупівель М1 і без навчання за власні кошти, без вилучень ФОП)
    grant_capex = GRANT; training = sum(x[1] for x in OWN if "Навчання" in x[0])
    opex_total = sum(x.opex_total for x in months) - grant_capex - training
    payroll_cost = sum(x.payroll_new + x.payroll_exist for x in months)
    esv_cost = sum(x.esv_emp for x in months)
    fop_taxes = sum(x.ep + x.vz_fop + x.esv_fop for x in months)
    op_costs = opex_total + payroll_cost + esv_cost + fop_taxes
    op_profit = rev - op_costs                       # до вилучень власника
    profitability = op_profit / op_costs
    # за роками
    years = []
    for y in range(3):
        ms = months[y*12:(y+1)*12]
        r = sum(x.revenue for x in ms)
        oc = (sum(x.opex_total for x in ms) - (grant_capex if y == 0 else 0) - sum(v for (n, v, mm) in OWN if "Навчання" in n and (mm-1)//12 == y)
              + sum(x.payroll_new + x.payroll_exist + x.esv_emp + x.ep + x.vz_fop + x.esv_fop for x in ms))
        years.append(dict(revenue=r, op_costs=oc, op_profit=r-oc, taxes=sum(x.taxes_total for x in ms),
                          payroll=sum(x.payroll_new+x.payroll_exist for x in ms)))
    invest_total = GRANT + OWN_TOTAL
    # окупність: інвестиції / середньорічний операційний прибуток
    avg_year_profit = op_profit / 3
    payback_years = invest_total / avg_year_profit if avg_year_profit > 0 else None
    # беззбитковість: перший місяць, коли виручка покриває місячні витрати (без грантових закупівель і навчання)
    be_month = None
    for x in months:
        monthly_costs = x.opex_total - (grant_capex if x.m == 1 else 0) - sum(v for (n, v, mm) in OWN if "Навчання" in n and mm == x.m) \
                        + x.payroll_new + x.payroll_exist + x.esv_emp + x.ep + x.vz_fop + x.esv_fop
        if x.revenue >= monthly_costs: be_month = x.m; break
    return dict(revenue=rev, taxes=taxes, taxes_from_m2=taxes_from_m2, cover_month=cover_month, cover_month_m2=cover_month_m2,
                inflow=inflow, outflow=outflow, end_balance=balances[-1], min_balance=min_bal, min_balance_month=min_bal_m,
                payroll_new_total=payroll_new_total, op_costs=op_costs, op_profit=op_profit, profitability=profitability,
                years=years, invest_total=invest_total, payback_years=payback_years, be_month=be_month,
                pdfo=sum(x.pdfo for x in months), vz_emp=sum(x.vz_emp for x in months), esv_emp=sum(x.esv_emp for x in months),
                ep=sum(x.ep for x in months), vz_fop=sum(x.vz_fop for x in months), esv_fop=sum(x.esv_fop for x in months),
                own_total=OWN_TOTAL, own_share=OWN_TOTAL/(GRANT+OWN_TOTAL), balances=balances,
                drawings=sum(x.drawings for x in months))

if __name__ == "__main__":
    s = summary()
    for k, v in s.items():
        if k in ("balances",): continue
        print(f"{k}: {v if not isinstance(v,float) else round(v,2)}")
    print("months:")
    for x in build_months():
        print(x.m, x.rev_main, x.rev_extra, x.revenue, "pay", x.payroll_new, "tax", x.taxes_total, "out", x.outflow, "in", x.inflow)

# --- Додаткові показники ефективності ---
def operating_cash_flows(months=None):
    """Операційний грошовий потік за місяцями (без грантових закупівель, навчання та вилучень ФОП)."""
    months = months or build_months()
    training_by_m = {mm: v for (n, v, mm) in OWN if "Навчання" in n}
    flows = []
    for x in months:
        capex = GRANT if x.m == 1 else 0
        costs = x.opex_total - capex - training_by_m.get(x.m, 0) + x.payroll_new + x.payroll_exist + x.esv_emp + x.ep + x.vz_fop + x.esv_fop
        flows.append(x.revenue - costs)
    return flows

def irr_annual(months=None):
    """IRR проєкту: інвестиції (грант + власні кошти) проти операційних потоків за 36 місяців."""
    flows = operating_cash_flows(months)
    cf = [-(GRANT + OWN_TOTAL)] + flows
    def npv(r):
        return sum(c / (1 + r) ** i for i, c in enumerate(cf))
    lo, hi = -0.99, 10.0
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0: hi = mid
        else: lo = mid
    monthly = (lo + hi) / 2
    return (1 + monthly) ** 12 - 1

def roi(months=None):
    s = summary(months)
    return s["op_profit"] / s["invest_total"]
