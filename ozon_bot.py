import os
import sys
import nest_asyncio
import pandas as pd
from io import BytesIO
from telegram import Update, InputFile
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, ContextTypes, filters

sku_costs = {
    "ФСО_MAX_2": 215,
    "ФСО_MAX_4": 415,
    "FARA_NIVA_2": 1060,
}

sku_names = {
    "ФСО_MAX_2": "📦 ФСО_MAX_2",
    "ФСО_MAX_4": "📦 ФСО_MAX_4",
    "FARA_NIVA_2": "📦 FARA_NIVA_2"
}

expense_labels = {
    "Бонусы продавца": "🎁 Бонусы продавца",
    "Вознаграждение за продажу": "💼 Вознаграждение за продажу",
    "Логистика": "🚚 Логистика",
    "Продвижение с оплатой за заказ": "📢 Продвижение (оплата за заказ)",
    "Эквайринг": "💳 Эквайринг"
}

def format_rub(value):
    return f"{value:,.2f} ₽".replace(",", " ")

def extract_report(file_path):
    df = pd.read_excel(file_path, header=None)
    for i in range(10):
        if "Тип начисления" in df.iloc[i].astype(str).values:
            df = pd.read_excel(file_path, header=i)
            break
    else:
        return "❌ Не найдена строка с заголовками."

    df.columns = df.columns.str.strip()
    df = df[df["Тип начисления"].str.lower() != " "]
    df["Количество"] = pd.to_numeric(df["Количество"], errors="coerce").fillna(0)
    df["Цена продавца"] = pd.to_numeric(df["Цена продавца"], errors="coerce").fillna(0)
    df["Сумма итого, руб"] = pd.to_numeric(df["Сумма итого, руб"], errors="coerce").fillna(0)

    sales = df[df["Тип начисления"].str.lower() == "выручка"].copy()
    sales["Сумма"] = sales["Количество"] * sales["Цена продавца"]
    total_revenue = sales["Сумма"].sum()

    sku_column = next((col for col in df.columns if any(k in col.lower() for k in ["артикул", "sku", "объявление"])), None)
    if not sku_column:
        return "❌ Не найдена колонка с артикулами товара."
    qty_by_sku = sales.groupby(sku_column)["Количество"].sum().to_dict()
    cost_by_sku = {sku: qty_by_sku.get(sku, 0) * cost for sku, cost in sku_costs.items()}
    total_cost = sum(cost_by_sku.values())

    expenses = df[~df["Тип начисления"].str.lower().isin(["выручка"])]
    expenses = expenses[~expenses["Тип начисления"].str.lower().str.contains("баллы")]
    expenses_sum = expenses.groupby("Тип начисления")["Сумма итого, руб"].sum().abs().to_dict()
    total_expenses = sum(expenses_sum.values())

    tax = total_revenue * 0.04
    net_profit = total_revenue - total_cost - total_expenses - tax
    gross_margin = (total_revenue - total_cost) / total_revenue * 100 if total_revenue else 0
    net_margin = net_profit / total_revenue * 100 if total_revenue else 0
    total_qty = sum(qty_by_sku.values())
    avg_check = total_revenue / total_qty if total_qty else 0

    filename = os.path.basename(file_path)
    lines = [f"📊 Финансовый отчёт на основе: {filename}\n"]
    lines.append(f"💰 Общая выручка: {format_rub(total_revenue)}")
    lines.append(f"🧾 Себестоимость: {format_rub(total_cost)}")
    for sku in sku_costs:
        sku_name = sku_names.get(sku, sku)
        lines.append(f"   ├─ {sku_name}: {format_rub(cost_by_sku.get(sku, 0))}")
    lines[-1] = lines[-1].replace("├─", "└─")

    lines.append(f"\n💸 Прочие расходы: {format_rub(total_expenses)}")
    keys = list(expenses_sum.keys())
    for i, key in enumerate(keys):
        label = expense_labels.get(key, f"▫️ {key}")
        prefix = "└─" if i == len(keys)-1 else "├─"
        lines.append(f"   {prefix} {label}: {format_rub(expenses_sum[key])}")

    lines += [
        f"\n📉 Налог по УСН (4%): {format_rub(tax)}",
        f"\n🟢 Чистая прибыль: {format_rub(net_profit)}\n",
        f"🧮 Количество продаж: {int(total_qty)} шт."
    ]
    for sku in sku_costs:
        sku_name = sku_names.get(sku, sku)
        lines.append(f"   ├─ {sku_name}: {int(qty_by_sku.get(sku, 0))} шт.")
    lines[-1] = lines[-1].replace("├─", "└─")
    lines += [
        f"\n💳 Средний чек: {format_rub(avg_check)}",
        f"📈 Валовая маржа: {gross_margin:.2f}%",
        f"📊 Чистая маржа: {net_margin:.2f}%"
    ]
    return "\n".join(lines)

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = update.message.document
    if not file.file_name.endswith(".xlsx"):
        await update.message.reply_text("⚠️ Принимаются только Excel .xlsx файлы.")
        return
    file_path = f"temp_{file.file_name}"
    new_file = await context.bot.get_file(file.file_id)
    await new_file.download_to_drive(file_path)
    text = extract_report(file_path)
    os.remove(file_path)
    await update.message.reply_text("📎 Ваш отчёт готов! Спасибо, что пользуетесь ботом 🙌")
    with BytesIO(text.encode("utf-8")) as f:
        f.name = "OZON_финансовый_отчет.txt"
        await update.message.reply_document(document=InputFile(f), caption="🧾 Финансовый отчёт 📊")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📎 Пришли Excel-файл от Ozon — я верну текстовый отчёт с расшифровкой.")

async def main():
    import nest_asyncio
    import asyncio
    nest_asyncio.apply()
    app = ApplicationBuilder().token(os.getenv("BOT_TOKEN")).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.MimeType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_file))
    await app.run_polling()

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
