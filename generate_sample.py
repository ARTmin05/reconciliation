"""Генерирует две выдуманные таблицы для демонстрации сверки:

- orders.xlsx — заказы из CRM интернет-магазина (что должно было быть оплачено);
- bank.xlsx   — выписка банка (какие деньги реально пришли).

В данные специально заложены типичные расхождения.
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

OUT = Path(__file__).parent / "input"
random.seed(11)

CLIENTS = ["ИП Смирнов А.В.", "ООО «Ромашка»", "Иванова Е.С.", "ООО «ТехноСнаб»", "Петров Д.И.",
           "ООО «Вектор»", "Кузнецова М.А.", "ИП Орлов К.П.", "ООО «Альфа Строй»", "Соколов Н.Н.",
           "ООО «Гарант»", "Морозова О.В.", "ИП Белов С.С.", "ООО «Меридиан»", "Волков А.А."]


def main():
    start = date.today() - timedelta(days=30)
    orders, bank = [], []

    for i in range(1, 181):
        order_id = 10400 + i
        day = start + timedelta(days=random.randint(0, 29))
        client = random.choice(CLIENTS)
        amount = random.choice([1490, 2990, 4500, 7800, 12500, 15900, 24000, 39900])
        orders.append({"Номер заказа": order_id, "Дата заказа": day, "Клиент": client,
                       "Сумма, ₽": amount, "Статус": "Оплачен"})

        roll = random.random()
        if roll < 0.05:      # деньги не пришли, хотя в CRM «Оплачен»
            continue
        pay_day = day + timedelta(days=random.choice([0, 0, 1, 1, 2]))
        pay_amount = amount
        purpose = f"Оплата по заказу №{order_id}"
        if roll < 0.10:      # пришла не та сумма (недоплата / переплата)
            pay_amount = amount - random.choice([100, 500, 990]) if random.random() < 0.7 else amount + 500
        elif roll < 0.14:    # плательщик не указал номер заказа
            purpose = random.choice(["Оплата", "За товар", "Оплата по счёту", "Перевод"])
        bank.append({"Дата": pay_day, "Плательщик": client, "Сумма, ₽": pay_amount,
                     "Назначение платежа": purpose})
        if roll > 0.985:     # двойная оплата
            bank.append({"Дата": pay_day + timedelta(days=1), "Плательщик": client,
                         "Сумма, ₽": pay_amount, "Назначение платежа": purpose})

    for _ in range(4):       # деньги пришли, а заказа в CRM нет
        bank.append({"Дата": start + timedelta(days=random.randint(0, 29)), "Плательщик": random.choice(CLIENTS),
                     "Сумма, ₽": random.choice([3500, 8900, 15000]),
                     "Назначение платежа": f"Оплата по заказу №{random.randint(20000, 20999)}"})

    random.shuffle(bank)
    OUT.mkdir(exist_ok=True)
    pd.DataFrame(orders).to_excel(OUT / "orders.xlsx", index=False)
    pd.DataFrame(bank).to_excel(OUT / "bank.xlsx", index=False)
    print(f"Заказов: {len(orders)}, платежей в выписке: {len(bank)} → {OUT}")


if __name__ == "__main__":
    main()
