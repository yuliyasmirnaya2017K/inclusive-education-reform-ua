import os
import pandas as pd

# 1. Читаємо локальний файл, який лежить поруч зі скриптом
input_file = "dataset_2026-11-05.xlsx"
output_file = "osvita_perfect_tableau.xlsx"

# Перевіряємо, чи файл на місці, або шукаємо схожий Excel-файл
if not os.path.exists(input_file):
    excel_files = [f for f in os.listdir(".") if f.endswith(".xlsx")]
    if excel_files:
        input_file = excel_files[0]
        print(f"📌 Знайдено файл із назвою: {input_file}")
    else:
        raise FileNotFoundError(
            f"❌ Помилка: Будь ласка, покладіть файл '{input_file}' у папку зі скриптом!"
        )

print(f"⏳ Обробка файлу '{input_file}'...")
df = pd.read_excel(input_file)

# Очищуємо назви колонок від зайвих пробілів
df.columns = df.columns.str.strip()

# 2. Автоматично визначаємо ключові колонки
col_name = (
    "Розріз"
    if "Розріз" in df.columns
    else ("Rozpis" if "Rozpis" in df.columns else df.columns[1])
)
geo_col = (
    "Територіальний розріз"
    if "Територіальний розріз" in df.columns
    else df.columns[0]
)

# Визначаємо роки для аналізу
years = [str(y) for y in range(2019, 2026) if str(y) in df.columns]

# 3. Фільтруємо за ключовими словами та робимо короткі назви категорій
df_filtered = df[
    df[col_name].str.contains("Інклюзивна|Спеціального", na=False)
].copy()
df_filtered[col_name] = df_filtered[col_name].apply(
    lambda x: "Спеціальна" if "Спеціального" in str(x) else "Інклюзивна"
)

# 4. Робимо Unpivot (Melt) — розгортаємо роки по вертикалі для Tableau
df_melted = pd.melt(
    df_filtered,
    id_vars=[geo_col, col_name],
    value_vars=years,
    var_name="Рік",
    value_name="Кількість груп",
)

# 5. Приводимо кількість груп до цілих чисел
df_melted["Кількість груп"] = (
    pd.to_numeric(df_melted["Кількість груп"], errors="coerce")
    .fillna(0)
    .astype(int)
)

# 6. Зберігаємо очищений результат у новий файл Excel
df_melted.to_excel(output_file, index=False)
print(
    f"🎉 Успіх! Очищений файл '{output_file}' успішно створено поруч зі скриптом!"
)
