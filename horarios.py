import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta, time
import pandas as pd
import matplotlib.pyplot as plt

class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Weekly Schedule Planner")
        self.entries = []
        self.setup_ui()

    def setup_ui(self):
        # Configuración de la interfaz
        input_frame = ttk.Frame(self.root, padding="10")
        input_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Selector de día
        ttk.Label(input_frame, text="Day:").grid(row=0, column=0, sticky=tk.W)
        self.day_var = tk.StringVar()
        self.day_cb = ttk.Combobox(input_frame, textvariable=self.day_var, 
                                  values=['Monday', 'Tuesday', 'Wednesday', 'Thursday', 
                                          'Friday', 'Saturday', 'Sunday'])
        self.day_cb.grid(row=0, column=1, sticky=tk.W)

        # Selectores de hora
        time_options = self.generate_time_options()
        ttk.Label(input_frame, text="Start Time:").grid(row=1, column=0, sticky=tk.W)
        self.start_var = tk.StringVar()
        self.start_cb = ttk.Combobox(input_frame, textvariable=self.start_var, values=time_options)
        self.start_cb.grid(row=1, column=1, sticky=tk.W)

        ttk.Label(input_frame, text="End Time:").grid(row=2, column=0, sticky=tk.W)
        self.end_var = tk.StringVar()
        self.end_cb = ttk.Combobox(input_frame, textvariable=self.end_var, values=time_options)
        self.end_cb.grid(row=2, column=1, sticky=tk.W)

        # Campo de actividad
        ttk.Label(input_frame, text="Activity:").grid(row=3, column=0, sticky=tk.W)
        self.activity_var = tk.StringVar()
        self.activity_entry = ttk.Entry(input_frame, textvariable=self.activity_var)
        self.activity_entry.grid(row=3, column=1, sticky=tk.W)

        # Botones de acción
        self.add_btn = ttk.Button(input_frame, text="Add", command=self.add_entry)
        self.add_btn.grid(row=4, column=0, pady=5)
        self.edit_btn = ttk.Button(input_frame, text="Edit", command=self.edit_entry, state=tk.DISABLED)
        self.edit_btn.grid(row=4, column=1, pady=5)
        self.delete_btn = ttk.Button(input_frame, text="Delete", command=self.delete_entry, state=tk.DISABLED)
        self.delete_btn.grid(row=4, column=2, pady=5)

        # Tabla de visualización
        tree_frame = ttk.Frame(self.root, padding="10")
        tree_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree = ttk.Treeview(tree_frame, columns=('Day', 'Start', 'End', 'Activity'), show='headings')
        self.tree.heading('Day', text='Day')
        self.tree.heading('Start', text='Start Time')
        self.tree.heading('End', text='End Time')
        self.tree.heading('Activity', text='Activity')
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # Botones de exportación
        export_frame = ttk.Frame(self.root, padding="10")
        export_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        ttk.Button(export_frame, text="Export to Image", command=self.export_image).grid(row=0, column=0, padx=5)
        ttk.Button(export_frame, text="Export to Excel", command=self.export_excel).grid(row=0, column=1, padx=5)

        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)

    def generate_time_options(self):
        times = []
        current = datetime.strptime('7:00 AM', '%I:%M %p')
        end = datetime.strptime('8:00 PM', '%I:%M %p')
        while current <= end:
            times.append(current.strftime('%I:%M %p'))
            current += timedelta(minutes=30)
        return times

    def add_entry(self):
        # Validación y adición de entradas
        day = self.day_var.get()
        start = self.start_var.get()
        end = self.end_var.get()
        activity = self.activity_var.get().strip()

        if not all([day, start, end, activity]):
            messagebox.showerror("Error", "All fields are required.")
            return

        try:
            start_time = datetime.strptime(start, '%I:%M %p').time()
            end_time = datetime.strptime(end, '%I:%M %p').time()
        except ValueError:
            messagebox.showerror("Error", "Invalid time format.")
            return

        if start_time >= end_time:
            messagebox.showerror("Error", "End time must be after start time.")
            return

        new_entry = {'day': day, 'start': start, 'end': end, 'activity': activity}
        if self.has_overlap(new_entry):
            messagebox.showerror("Error", "Time slot overlaps with existing entry.")
            return

        self.entries.append(new_entry)
        self.update_tree()
        self.clear_inputs()

    def has_overlap(self, new_entry, exclude_index=None):
        # Comprobación de superposición de horarios
        for i, entry in enumerate(self.entries):
            if exclude_index is not None and i == exclude_index:
                continue
            if entry['day'] != new_entry['day']:
                continue
            existing_start = datetime.strptime(entry['start'], '%I:%M %p').time()
            existing_end = datetime.strptime(entry['end'], '%I:%M %p').time()
            new_start = datetime.strptime(new_entry['start'], '%I:%M %p').time()
            new_end = datetime.strptime(new_entry['end'], '%I:%M %p').time()
            if (new_start < existing_end) and (new_end > existing_start):
                return True
        return False

    def update_tree(self):
        # Actualización de la tabla visual
        self.tree.delete(*self.tree.get_children())
        for entry in self.entries:
            self.tree.insert('', 'end', values=(entry['day'], entry['start'], entry['end'], entry['activity']))

    def on_tree_select(self, event):
        # Manejo de selección de elementos
        selected = self.tree.selection()
        if selected:
            self.edit_btn.config(state=tk.NORMAL)
            self.delete_btn.config(state=tk.NORMAL)
            item = self.tree.item(selected[0])
            self.day_var.set(item['values'][0])
            self.start_var.set(item['values'][1])
            self.end_var.set(item['values'][2])
            self.activity_var.set(item['values'][3])
        else:
            self.edit_btn.config(state=tk.DISABLED)
            self.delete_btn.config(state=tk.DISABLED)

    def edit_entry(self):
        # Edición de entradas existentes
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        old_values = item['values']
        index = next(i for i, e in enumerate(self.entries) if (e['day'], e['start'], e['end'], e['activity']) == tuple(old_values))

        new_entry = {
            'day': self.day_var.get(),
            'start': self.start_var.get(),
            'end': self.end_var.get(),
            'activity': self.activity_var.get().strip()
        }

        if self.has_overlap(new_entry, exclude_index=index):
            messagebox.showerror("Error", "Time slot overlaps with existing entry.")
            return

        self.entries[index] = new_entry
        self.update_tree()
        self.clear_inputs()

    def delete_entry(self):
        # Eliminación de entradas
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item['values']
        self.entries = [e for e in self.entries if (e['day'], e['start'], e['end'], e['activity']) != tuple(values)]
        self.update_tree()
        self.clear_inputs()

    def generate_grid_data(self):
        # Generación de datos para exportación
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        time_slots = []
        current = datetime.strptime('7:00 AM', '%I:%M %p')
        end_time = datetime.strptime('8:00 PM', '%I:%M %p')
        while current < end_time:
            slot_start = current.strftime('%I:%M %p')
            current += timedelta(minutes=30)
            slot_end = current.strftime('%I:%M %p')
            time_slots.append(f"{slot_start} - {slot_end}")

        grid_data = []
        for slot in time_slots:
            row = [slot]
            for day in days:
                activity = self.get_activity_for_slot(day, slot)
                row.append(activity)
            grid_data.append(row)
        return grid_data, ['Time Slot'] + days

    def get_activity_for_slot(self, day, slot):
        # Obtención de actividad para franja horaria
        start_str, end_str = slot.split(' - ')
        slot_start = datetime.strptime(start_str.strip(), '%I:%M %p').time()
        slot_end = datetime.strptime(end_str.strip(), '%I:%M %p').time()
        for entry in self.entries:
            if entry['day'] != day:
                continue
            entry_start = datetime.strptime(entry['start'], '%I:%M %p').time()
            entry_end = datetime.strptime(entry['end'], '%I:%M %p').time()
            if entry_start <= slot_start and entry_end >= slot_end:
                return entry['activity']
        return ''

    def export_image(self):
        # Exportación a imagen
        grid_data, columns = self.generate_grid_data()
        if not grid_data:
            messagebox.showwarning("Warning", "No data to export.")
            return

        plt.figure(figsize=(12, len(grid_data)*0.5))
        plt.axis('off')
        table = plt.table(cellText=grid_data, colLabels=columns, loc='center', cellLoc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1, 1.5)
        plt.savefig('schedule.png', bbox_inches='tight')
        plt.close()
        messagebox.showinfo("Success", "Exported to schedule.png")

    def export_excel(self):
        # Exportación a Excel de las entradas individuales
        if not self.entries:
            messagebox.showwarning("Warning", "No data to export.")
            return

        df = pd.DataFrame(self.entries)
        # Ordenar columnas para que salgan en el orden esperado
        df = df[['day', 'start', 'end', 'activity']]
        df.columns = ['Day', 'Start Time', 'End Time', 'Activity']
        try:
            df.to_excel('schedule.xlsx', index=False)
            messagebox.showinfo("Success", "Exported to schedule.xlsx")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    def clear_inputs(self):
        # Limpieza de campos
        self.day_var.set('')
        self.start_var.set('')
        self.end_var.set('')
        self.activity_var.set('')
        self.edit_btn.config(state=tk.DISABLED)
        self.delete_btn.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()