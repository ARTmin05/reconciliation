# Reconciliation — orders vs bank statement

Compares orders from a CRM / online store with a bank statement and shows exactly where the money doesn't add up: unpaid orders, wrong amounts, double payments and payments without an order. Works even when payers forget to put the order number in the payment purpose.

All data in this repo is **synthetic**.

![Python](https://img.shields.io/badge/Python-3.13-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoftexcel&logoColor=white)

## ✨ What it finds

| Status | Meaning |
|---|---|
| ✅ Matched | order and payment found, amounts equal |
| ✅ Matched without order number | payer didn't specify the order — matched by payer, amount and date (±3 days) |
| ❌ No payment in statement | the order is marked "paid" in CRM, but no money arrived |
| ⚠️ Amount mismatch | underpayment or overpayment, with the exact difference |
| ⚠️ Double payment | the same order was paid twice |
| 🔎 Payment without order | money arrived for an order number that doesn't exist in CRM |
| ❓ Not matched | payment with no order number and no matching order |

The order number is extracted from free-text payment purposes like *"Оплата по заказу №10452"*.

## 📊 Excel report

- **Итоги (Summary)** — totals, expected vs received money, the difference, counts per status
- **Расхождения (Discrepancies)** — only the problems, color-coded, with comments ("Underpayment 500 ₽", "Paid 2 times")
- **Заказы / Выписка** — both source tables with a reconciliation status for every row

## 📸 Screenshots

<!-- screenshots -->

## 🚀 Run

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python generate_sample.py                    # synthetic orders.xlsx and bank.xlsx in input/
python reconcile.py                          # → output/sverka_<date>.xlsx
python reconcile.py orders.xlsx bank.xlsx    # your own files
```

## 🔄 Adapting

The same approach reconciles any two tables: marketplace payouts vs orders, warehouse stock vs accounting, payment provider reports vs the platform's back office. Matching rules (keys, date window, amount tolerance) are set in one place in `reconcile.py`.

---

## 🇷🇺 Кратко по-русски

Сверка заказов из CRM с банковской выпиской: находит неоплаченные заказы, недоплаты и переплаты, двойные оплаты и платежи без заказа. Умеет сопоставлять платежи, в которых не указан номер заказа, — по плательщику, сумме и дате. Результат — Excel-отчёт с итогами и подсвеченными расхождениями.

**Нужно сверить две таблицы?** Пишите — [GitHub @ARTmin05](https://github.com/ARTmin05).
