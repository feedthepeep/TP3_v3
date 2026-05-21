import pandas as pd
import os;
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class SimpleExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Table Viewer")
        self.root.geometry("1300x700")
        self.root.resizable(False, False)
        
        main_container = ttk.Frame(root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Таблица 
        left_frame = ttk.Frame(main_container, width=1000, height=900)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        left_frame.pack_propagate(False) 
        
        # Панель управления
        right_frame = ttk.Frame(main_container, width=450, height=900, relief=tk.RIDGE, borderwidth=2)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH)
        right_frame.pack_propagate(False)  
        
        # Заголовок панели
        right_label = ttk.Label(right_frame, text="Панель управления", font=('Arial', 12, 'bold'))
        right_label.pack(pady=10)
        
        self.control_panel = ttk.Frame(right_frame)
        self.control_panel.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Верхняя панель
        top_frame = ttk.Frame(left_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.filename_label = ttk.Label(top_frame, text="", font=('Arial', 11, 'italic'), foreground='black')
        self.filename_label.pack(side=tk.RIGHT, padx=10)
        
        ttk.Button(top_frame, text="Открыть Excel", command=self.load_excel).pack(side=tk.LEFT, padx=5)
        
        self.info_label = ttk.Label(top_frame, text="Готов к загрузке")
        self.info_label.pack(side=tk.LEFT, padx=10)
        
        # Фрейм для таблицы
        table_frame = ttk.Frame(left_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Скроллы
        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree = ttk.Treeview(table_frame, yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        self.df = None
        
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                # Загружаем данные
                self.df = pd.read_excel(file_path, header=1)
                
                file_name = os.path.basename(file_path)
                self.filename_label.config(text=f"{file_name}")
                
                column_names = [
                    'Регион', 
                    'Код Региона',
                    '2011', '2012', '2013', '2014', '2015', '2016', '2017'
                ]
                self.df.columns = column_names
                
                self.df['Регион'] = self.df['Регион'].str.strip()
                
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                columns = [str(col) for col in self.df.columns]
                self.tree["columns"] = columns
                self.tree["show"] = "headings"
                
                # Установка фиксированной ширины колонок
                for i, col in enumerate(columns):
                    self.tree.heading(col, text=col)
                    if i == 0:  # Первая колонка (Регион)
                        self.tree.column(col, width=350, minwidth=350, anchor='w')
                    elif i == 1:  # Вторая колонка (Код Региона)
                        self.tree.column(col, width=120, minwidth=120, anchor='center')
                    else:  # Колонки с годами (2011-2017)
                        self.tree.column(col, width=100, minwidth=100, anchor='center')
                
                # Вставка данных
                for _, row in self.df.iterrows():
                    values = [str(val) if pd.notna(val) else "" for val in row]
                    self.tree.insert("", "end", values=values)
                
                self.info_label.config(text=f"Загружено: {len(self.df)} строк, {len(self.df.columns)} колонок")
            
                self.update_control_panel()
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{str(e)}")
    
    def update_control_panel(self):
        for widget in self.control_panel.winfo_children():
            widget.destroy()
        
        if self.df is not None:
            # Показываем информацию о данных
            info_text = f"""Информация о данных:
                            
                        Регионов: {len(self.df)}
                        Период: 2011-2017 гг.
                        Всего записей: {len(self.df) * 7}
                        """
            
            ttk.Label(self.control_panel, text=info_text, justify=tk.LEFT, 
                     font=('Arial', 10)).pack(pady=10, anchor='w')
            
            ttk.Separator(self.control_panel, orient='horizontal').pack(fill=tk.X, pady=10)
            
            ttk.Label(self.control_panel, text="Доступные действия:", 
                     font=('Arial', 10, 'bold')).pack(anchor='w', pady=(10, 5))
            
            ttk.Label(self.control_panel, text="Выберите регион:", 
                     font=('Arial', 9)).pack(anchor='w', pady=(10, 5))
            
            # Создание Listbox
            listbox_frame = ttk.Frame(self.control_panel)
            listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            
            # Список регионов
            self.region_listbox = tk.Listbox(listbox_frame, height=10, font=('Arial', 9))
            self.region_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Добавляем скролл
            listbox_scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.region_listbox.yview)
            listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.region_listbox.config(yscrollcommand=listbox_scrollbar.set)
            
            for region in self.df['Регион'].values:
                self.region_listbox.insert(tk.END, region)
            
            # Кнопка для построения графика
            self.plot_button = ttk.Button(self.control_panel, text="Построить график", 
                                         command=self.plot_migration, state='normal')
            self.plot_button.pack(fill=tk.X, pady=(10, 5))
            
        else:
            ttk.Label(self.control_panel, text="Загрузите Excel файл\nдля отображения информации", 
                     justify=tk.CENTER).pack(expand=True)
    
    def plot_migration(self):
        if self.df is None:
            messagebox.showwarning("Предупреждение", "Сначала загрузите Excel файл")
            return
        
        selection = self.region_listbox.curselection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите регион из списка")
            return
        
        selected_region = self.region_listbox.get(selection[0])
    
        region_data = self.df[self.df['Регион'] == selected_region]
        
        if region_data.empty:
            messagebox.showerror("Ошибка", "Данные для выбранного региона не найдены")
            return
        
        years = ['2011', '2012', '2013', '2014', '2015', '2016', '2017']
        values = [region_data[year].values[0] for year in years]
        
        # Создаем новое окно для графика
        plot_window = tk.Toplevel(self.root)
        plot_window.title(f"Миграция: {selected_region}")
        plot_window.geometry("900x700")
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Строим график
        ax.plot(years, values, marker='o', linewidth=2, markersize=8, color='blue')
        
        # Настройка вида графика
        ax.set_xlabel('Год', fontsize=12, fontweight='bold')
        ax.set_ylabel('Число выбывших', fontsize=12, fontweight='bold')
        ax.set_title(f'Динамика миграции: {selected_region}', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3, linestyle='--')
        plt.xticks(rotation=45)
        
        # Добавляем значения на график
        for i, (year, value) in enumerate(zip(years, values)):
            ax.annotate(f'{value:,.0f}', (year, value), textcoords="offset points", 
                       xytext=(0, 10), ha='center', fontsize=9)
        
        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=plot_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleExcelApp(root)
    root.mainloop()