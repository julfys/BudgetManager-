import openpyxl
from openpyxl.styles import Font
import tkinter as tk
from tkinter import messagebox, filedialog

def create_budget_table(total_budget, needs):
    # Расчёт остатка средств
    total_cost = sum(cost for _, cost in needs)
    remaining_budget = total_budget - total_cost

    # Создание Excel-файла
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Бюджет"

    # Заголовки
    sheet["A1"] = "Название"
    sheet["B1"] = "Стоимость"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"].font = Font(bold=True)

    # Заполнение данных о нуждах
    for row, (name, cost) in enumerate(needs, start=2):
        sheet[f"A{row}"] = name
        sheet[f"B{row}"] = cost

    # Итоговые значения
    summary_row = len(needs) + 2
    sheet[f"A{summary_row}"] = "Итого:"
    sheet[f"B{summary_row}"] = total_cost
    sheet[f"A{summary_row}"].font = Font(bold=True)
    sheet[f"B{summary_row}"].font = Font(bold=True)

    remaining_row = summary_row + 1
    sheet[f"A{remaining_row}"] = "Остаток бюджета:"
    sheet[f"B{remaining_row}"] = remaining_budget
    sheet[f"A{remaining_row}"].font = Font(bold=True)
    sheet[f"B{remaining_row}"].font = Font(bold=True)

    # Сохранение файла
    file_name = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_name:
        wb.save(file_name)
        messagebox.showinfo("Успех", f"Ваш бюджет сохранён в файл '{file_name}'.")


def add_need():
    name = need_name_entry.get()
    try:
        cost = float(need_cost_entry.get())
    except ValueError:
        messagebox.showerror("Ошибка", "Введите корректную стоимость.")
        return

    if name and cost >= 0:
        needs.append((name, cost))
        needs_list.insert(tk.END, f"{name}: {cost} руб.")
        need_name_entry.delete(0, tk.END)
        need_cost_entry.delete(0, tk.END)
    else:
        messagebox.showerror("Ошибка", "Введите название нужды и корректную стоимость.")

def save_budget():
    try:
        total_budget = float(budget_entry.get())
    except ValueError:
        messagebox.showerror("Ошибка", "Введите корректную сумму бюджета.")
        return

    if needs:
        create_budget_table(total_budget, needs)
    else:
        messagebox.showerror("Ошибка", "Добавьте хотя бы одну нужду.")

# Цветовая гамма
COLORS = {
    "background": "#fdf0d5",
    "header": "#780000",
    "button": "#c1121f",
    "text": "#003049",
    "entry": "#669bbc",
}

# Интерфейс программы
root = tk.Tk()
root.title("Управление бюджетом")
root.configure(bg=COLORS["background"])

# Заголовок
header = tk.Label(root, text="Программа управления бюджетом", font=("Arial", 16, "bold"), bg=COLORS["header"], fg="white")
header.pack(pady=10, fill=tk.X)

# Ввод бюджета
budget_frame = tk.Frame(root, bg=COLORS["background"])
budget_frame.pack(pady=5)

budget_label = tk.Label(budget_frame, text="Общий бюджет:", font=("Arial", 12), bg=COLORS["background"], fg=COLORS["text"])
budget_label.pack(side=tk.LEFT, padx=5)

budget_entry = tk.Entry(budget_frame, font=("Arial", 12), bg=COLORS["entry"], fg="black")
budget_entry.pack(side=tk.LEFT, padx=5)

# Ввод нужд
needs_frame = tk.Frame(root, bg=COLORS["background"])
needs_frame.pack(pady=5)

need_name_label = tk.Label(needs_frame, text="Название нужды:", font=("Arial", 12), bg=COLORS["background"], fg=COLORS["text"])
need_name_label.grid(row=0, column=0, padx=5, pady=5)

need_name_entry = tk.Entry(needs_frame, font=("Arial", 12), bg=COLORS["entry"], fg="black")
need_name_entry.grid(row=0, column=1, padx=5, pady=5)

need_cost_label = tk.Label(needs_frame, text="Стоимость:", font=("Arial", 12), bg=COLORS["background"], fg=COLORS["text"])
need_cost_label.grid(row=1, column=0, padx=5, pady=5)

need_cost_entry = tk.Entry(needs_frame, font=("Arial", 12), bg=COLORS["entry"], fg="black")
need_cost_entry.grid(row=1, column=1, padx=5, pady=5)

add_need_button = tk.Button(needs_frame, text="Добавить", font=("Arial", 12), bg=COLORS["button"], fg="white", command=add_need)
add_need_button.grid(row=0, column=2, rowspan=2, padx=5, pady=5)

# Список нужд
needs_list = tk.Listbox(root, font=("Arial", 12), bg=COLORS["entry"], fg="black", height=10, width=50)
needs_list.pack(pady=10)

# Сохранение бюджета
save_button = tk.Button(root, text="Сохранить бюджет", font=("Arial", 12), bg=COLORS["button"], fg="white", command=save_budget)
save_button.pack(pady=10)

# Данные о нуждах
needs = []

root.mainloop()