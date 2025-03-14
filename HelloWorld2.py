import pandas as pd
import os

products = pd.read_excel(r"C:\Users\user\Desktop\products.xlsx")
orders = pd.read_excel(r"C:\Users\user\Desktop\orders.xlsx")

df = orders.merge(products, on="product_id")

# Задача 1 - Самая ходовая товарная группа
category_sales = df.groupby("level1")["quantity"].sum().reset_index()
category_sales = category_sales.rename(columns={"level1": "Категория", "quantity": "Количество"})
category_sales = category_sales.sort_values(by="Количество", ascending=False)
print(category_sales)

category_sales.to_excel(r"C:\Users\user\Desktop\category_sales.xlsx", index=False)

import matplotlib.pyplot as plt

plt.figure(figsize=(14, 8))
plt.barh(category_sales["Категория"], category_sales["Количество"], color="skyblue")
plt.xlabel("Количество проданных товаров")
plt.ylabel("Категория товаров")
plt.title("Самая ходовая товарная группа")
plt.yticks(fontsize=7)
plt.gca().invert_yaxis()  # чтобы самая популярная категория была сверху
plt.show()

# Задача 2 - Распределение продаж по подкатегориям
subcategory_sales = df.groupby(["level1", "level2"])["quantity"].sum().reset_index()
subcategory_sales = subcategory_sales.rename(columns={"level1": "Категория", "level2": "Подкатегория", "quantity": "Количество"})
subcategory_sales.to_excel(r"C:\Users\user\Desktop\subcategory_sales.xlsx", index=False)
print(subcategory_sales)

# plotdata = pd.DataFrame({
#    "Все для суши":[9],
#    "Зерновые для завтраков":[24],
#    "Ингредиенты для готовки":[21],
#    "Крупы, бобовые":[30],
#    "Макаронные изделия":[24]
#    }, 
#    index=["Бакалея"]
#)
#plotdata.plot(kind='bar', stacked=True)
#plt.title("Тест")
#plt.xlabel("Категория")
#plt.ylabel("Продажи")
#plt.legend(title="Категория", bbox_to_anchor=(1.05, 1), loc='upper left')
#plt.tight_layout()
#plt.show()

num_rows = 25  # чтобы выводить 25 строк на одном графике
total_rows = len(subcategory_sales)

for start in range(0, total_rows, num_rows):
    end = min(start + num_rows, total_rows)
    sub_table = subcategory_sales.iloc[start:end]

    fig, ax = plt.subplots(figsize=(12, 6))
    ax.axis("tight")
    ax.axis("off")

    table = ax.table(cellText=sub_table.values, 
                     colLabels=sub_table.columns, 
                     rowLabels=sub_table.index, 
                     cellLoc="center", 
                     loc="center")

    plt.title(f"Распределение продаж (строки {start + 1} - {end})")
    plt.show()

# Задача 3 - Найти средний чек в заданную дату
df["accepted_at"] = pd.to_datetime(df["accepted_at"]).dt.date

target_date = pd.to_datetime("2022-01-13").date() # фильтруем заказы за 13.01.2022
daily_orders = df[df["accepted_at"] == target_date]

average_check = daily_orders.groupby("order_id")["price"].sum().mean() # считаем сумму по каждому чеку

print(f"Средний чек за {target_date}: {average_check:.2f}")

fig, ax = plt.subplots(figsize=(6, 3)) # создаем визуализацию
ax.axis("off")  # убираем оси
text = f"Средний чек за {target_date}:\n{average_check:.2f} руб."
ax.text(0.5, 0.5, text, fontsize=14, ha="center", va="center", fontweight="bold")

plt.show()

# Задача 4 - Доля промо в заданной категории

cheese_sales = df[df["level1"] == "Сыры"] # фильтруем данные по категории "Сыры"

promo_sales = cheese_sales[cheese_sales["regular_price"] != cheese_sales["price"]] # количество товаров, которые были проданы по промо (где цена отличается от базовой)
regular_sales = cheese_sales[cheese_sales["regular_price"] == cheese_sales["price"]]

promo_count = promo_sales["quantity"].sum()
regular_count = regular_sales["quantity"].sum()

promo_share = promo_count / (promo_count + regular_count) * 100 # рассчитываем долю промо
regular_share = 100 - promo_share

labels = ["Промо", "Обычные продажи"] # данные для pie chart
sizes = [promo_share, regular_share]
colors = ['gold', 'lightgray']

plt.figure(figsize=(7, 7)) # строим pie chart
plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 10})
plt.title("Доля промо в продажах категории 'Сыры'", fontsize=14)
plt.show()

print(f"Доля промо-продаж в категории 'Сыры': {promo_share:.2f}%")

# Задача 5 - Посчитать маржу по категориям
df["margin_rub"] = (df["price"] - df["cost_price"]) * df["quantity"] # рассчитываем маржу в рублях и процентах
df["total_sales"] = df["price"] * df["quantity"]
df["margin_percent"] = (df["margin_rub"] / df["total_sales"]) * 100

category_margin = df.groupby("level1").agg(
    margin_rub=("margin_rub", "sum"),
    margin_percent=("margin_percent", "mean")
).reset_index() # группируем данные по категориям и считаем сумму маржи и суммарные продажи

category_margin_sorted_rub = category_margin.sort_values(by="margin_rub", ascending=False) # сортируем по марже в рублях и процентах
category_margin_sorted_percent = category_margin.sort_values(by="margin_percent", ascending=False)

plt.figure(figsize=(14, 8)) # строим визуализацию

# bar chart для маржи в рублях
plt.subplot(1, 2, 1)
plt.barh(category_margin_sorted_rub["level1"], category_margin_sorted_rub["margin_rub"], color="skyblue")
plt.xlabel("Маржа (руб.)")
plt.ylabel("Категория товаров")
plt.title("Маржа по категориям (в рублях)")
plt.gca().invert_yaxis()  # чтобы самая высокая маржа была сверху

# bar chart для маржи в процентах
plt.subplot(1, 2, 2)
plt.barh(category_margin_sorted_percent["level1"], category_margin_sorted_percent["margin_percent"], color="lightgreen")
plt.xlabel("Маржа (%)")
plt.ylabel("Категория товаров")
plt.title("Маржа по категориям (в %)")
plt.gca().invert_yaxis()  # чтобы самая высокая маржа была сверху

plt.tight_layout()
plt.show()

print(category_margin) # выводим результат

# Задача 6 - ABC анализ
df["total_sales"] = df["price"] * df["quantity"] # вычисляем количество и сумму продаж

category_sales = df.groupby("level1").agg(
    total_quantity=("quantity", "sum"),
    total_sales=("total_sales", "sum")
).reset_index() # группируем по категориям

category_sales["Количество_ранг"] = category_sales["total_quantity"].rank(pct=True) # ABC-анализ по количеству
category_sales["Количество_группа"] = category_sales["Количество_ранг"].apply(
    lambda x: "A" if x <= 0.5 else ("B" if x <= 0.8 else "C")
)

category_sales["Продажи_ранг"] = category_sales["total_sales"].rank(pct=True) # ABC-анализ по сумме продаж
category_sales["Продажи_группа"] = category_sales["Продажи_ранг"].apply(
    lambda x: "A" if x <= 0.5 else ("B" if x <= 0.8 else "C")
)

category_sales["Итоговая группа"] = category_sales["Количество_группа"] + " " + category_sales["Продажи_группа"] # итоговая группа

category_sales = category_sales[["level1", "Количество_группа", "Продажи_группа", "Итоговая группа"]] # чтобы выводились только нужные столбцы

category_sales.rename(columns={"level1": "Категория"}, inplace=True) # переименовываем level1 в категория

output_file_path = r"C:\Users\user\Desktop\category_sales.xlsx" # сохраняем итоговую таблицу
category_sales.to_excel(output_file_path, index=False)

print(category_sales) # выводим результат

fig, ax = plt.subplots(figsize=(12, 6)) # строим визуализацию
ax.axis("tight")
ax.axis("off")

table = ax.table(cellText=category_sales.values, 
                 colLabels=category_sales.columns, 
                 cellLoc="center", 
                 loc="center") # отображаем таблицу

plt.title("ABC-анализ по категориям", pad=20) # добавляем заголовок

plt.show()