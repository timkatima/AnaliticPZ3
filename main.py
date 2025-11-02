import pandas as pd
import numpy as np
from scipy.stats import pointbiserialr
from datetime import datetime

# 1. Загрузка данных
try:
    orders_df = pd.read_csv('orders.csv', sep=';')
    cars_df = pd.read_excel('datasetNew.xlsx')
except FileNotFoundError:
    print("Файл 'orders.csv' или 'datasetNew.xlsx' не найден.")
    exit()
except Exception as e:
    print(f"Ошибка при загрузке файлов: {e}")
    exit()

# Удаляем лишний столбец, если есть
if '472' in orders_df.columns:
    orders_df = orders_df.drop(columns=['472'])

# 2. Проверка столбцов в orders.csv
required_order_columns = ['UserID', 'TotalPrice', 'Status', 'Product', 'Discount']
missing_order_columns = [col for col in required_order_columns if col not in orders_df.columns]
if missing_order_columns:
    print(f"Ошибка: В orders.csv отсутствуют столбцы: {missing_order_columns}")
    print(f"Доступные столбцы: {list(orders_df.columns)}")
    exit()

# Проверка столбцов в datasetNew.xlsx
required_car_columns = ['make', 'body']
missing_car_columns = [col for col in required_car_columns if col not in cars_df.columns]
if missing_car_columns:
    print(f"Ошибка: В datasetNew.xlsx отсутствуют столбцы: {missing_car_columns}")
    exit()

# Удаляем аномальные продукты (например, даты в Product)
orders_df = orders_df[~orders_df['Product'].str.contains(r'\d{4}-\d{2}-\d{2}', na=False)]

# 3. Средний доход с пользователя (в BYN)
non_canceled_orders = orders_df[orders_df['Status'].isin(['Paid', 'Delivered'])]
user_revenue = non_canceled_orders.groupby('UserID')['TotalPrice'].sum()
avg_revenue_per_user = user_revenue.mean()
print(f"Средний доход с пользователя: {avg_revenue_per_user:.2f} BYN")

# 4. Медианное значение возврата на одного пользователя (в BYN)
canceled_orders = orders_df[orders_df['Status'].isin(['Canceled_Mismatch', 'Canceled_Error'])]
user_returns = canceled_orders.groupby('UserID')['TotalPrice'].sum()
median_returns_per_user = user_returns.median()
median_returns_per_user = median_returns_per_user if not pd.isna(median_returns_per_user) else 0
print(f"Медианное значение возврата на одного пользователя: {median_returns_per_user:.2f} BYN")

# 5. Артикул товара, приносящего наибольший доход
product_revenue = non_canceled_orders.groupby('Product')['TotalPrice'].sum()
top_product = product_revenue.idxmax()
top_product_revenue = product_revenue.max()
print(f"Товар, приносящий наибольший доход: {top_product}, доход: {top_product_revenue:.2f} BYN")

# 6. Процент отмен для групп товаров (по бренду make)
orders_df['Product'] = orders_df['Product'].astype(str)
cars_df['Product'] = (cars_df['make'].astype(str) + ' ' +
                      cars_df['model'].astype(str) + ' ' +
                      cars_df['trim'].astype(str))
orders_with_make = orders_df.merge(cars_df[['Product', 'make']], on='Product', how='left')
orders_with_make['make'] = orders_with_make['make'].fillna('Unknown')

# Выбираем три популярные группы по make
top_makes = orders_with_make['make'].value_counts().head(5).index
group_cancellation_rates = {}
for make in top_makes:
    make_orders = orders_with_make[orders_with_make['make'] == make]
    total_orders = len(make_orders)
    canceled_orders = len(make_orders[make_orders['Status'].isin(['Canceled_Mismatch', 'Canceled_Error'])])
    cancellation_rate = (canceled_orders / total_orders) * 100 if total_orders > 0 else 0
    group_cancellation_rates[make] = cancellation_rate
print("Процент отмен по группам товаров (по бренду):")
for make, rate in group_cancellation_rates.items():
    print(f"{make}: {rate:.2f}%")

# 7. Зависимость между ошибочными заказами и размером скидки
orders_df['IsError'] = (orders_df['Status'] == 'Canceled_Error').astype(int)
correlation, p_value = pointbiserialr(orders_df['IsError'], orders_df['Discount'])
print(f"Зависимость между ошибочными заказами и скидкой:")
print(f"Коэффициент корреляции (Point-Biserial): {correlation:.4f}")
print(f"p-value: {p_value:.4f}")
if p_value < 0.05:
    print("Зависимость статистически значима (p < 0.05).")
else:
    print("Зависимость не является статистически значимой (p >= 0.05).")

# 8. Сохранение результатов
results = pd.DataFrame({
    'Metric': [
        'Средний доход с пользователя (BYN)',
        'Медианное значение возврата на пользователя (BYN)',
        'Товар с наибольшим доходом',
        'Доход от товара (BYN)'
    ],
    'Value': [
        f"{avg_revenue_per_user:.2f}",
        f"{median_returns_per_user:.2f}",
        top_product,
        f"{top_product_revenue:.2f}"
    ]
})
results = pd.concat([results, pd.DataFrame({
    'Metric': [f"Процент отмен для {make} (%)" for make in group_cancellation_rates],
    'Value': [f"{rate:.2f}" for rate in group_cancellation_rates.values()]
})], ignore_index=True)
results = pd.concat([results, pd.DataFrame({
    'Metric': ['Коэффициент корреляции (ошибочные заказы и скидка)', 'p-value'],
    'Value': [f"{correlation:.4f}", f"{p_value:.4f}"]
})], ignore_index=True)
results.to_csv('analytics_results.csv', index=False)
print("\nРезультаты сохранены в 'analytics_results.csv'")