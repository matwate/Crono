import openpyxl as ox
import customtkinter as ctk
import subprocess
import os
"""
What this code will do is open a window with 9 input fields.
one for each worker, with a Turno, Codigo and Name fields.

User will input the year of the Cronogram we're making and it will generate
the whole year's cronogram in a .xlsx file.

(At some point i will make it modular and API so i can 
host it on a server for my dad to not have to install anything 
and make him pay for my domain name and hosting)

"""

# Imma be doing the UI first
# Each field will be 3 input fields, Turno \ Codigo \ Name


class worker_fields(ctk.CTkFrame):
    def __init__(self, master, n, **kwargs):
        super().__init__(master, **kwargs)
        self.fields = [
            ctk.CTkEntry(self, placeholder_text="Turno"),
            ctk.CTkEntry(self, placeholder_text="Codigo"),
            ctk.CTkEntry(self, placeholder_text="Nombre"),
        ]
        self.label = ctk.CTkLabel(self, text=f"Trabajador {n}")
        self.label.grid(row=0, column=0, padx=10)
        self.fields[0].grid(row=0, column=1, padx=10)
        self.fields[1].grid(row=0, column=2, padx=10)
        self.fields[2].grid(row=0, column=3, padx=10)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("1000x800")
        self.workers = []
        for i in range(9):
            self.workers.append(worker_fields(self, i + 1))
            self.workers[i].grid(row=i, column=0, sticky="nsew", pady=20)
        self.day = None

        def whichday_callback(choice):
            self.day = choice

        self.whichday = ctk.CTkOptionMenu(
            self,
            values=[
                "Lunes",
                "Martes",
                "Miercoles",
                "Jueves",
                "Viernes",
                "Sabado",
                "Domingo",
            ],
            command=whichday_callback,
        )
        self.whichday.grid(row=10, column=1, pady=20)
        self.generate = ctk.CTkButton(
            self, text="Generar Cronograma", command=self.generate_cronogram
        )
        self.generate.grid(row=11, column=0, pady=20)

    def generate_cronogram(self):
        """
        Generates a yearly work schedule with the following rules:
        First block (workers 0-4):
        - Mon-Fri: 3 workers have D, 1 has N, 1 has A
        - Sat-Sun: D workers get L, N and A workers get TC
        - N and A roles rotate between the 5 workers every Tuesday

        Second block (workers 5-8):
        - Mon-Fri: All workers have D
        - Odd weeks: Workers 1&4 have TC on weekends, 2&3 have L
        - Even weeks: Workers 1&4 have L on weekends, 2&3 have TC
        """
        print("Generating cronogram...")
        self.day = self.whichday.get()
        
        # Create workbook
        wb = ox.Workbook()
        ws = wb.active
        ws.title = "Cronograma"
        
        # Get worker data
        workers = []
        for worker_field in self.workers:
            workers.append({
                'turno': worker_field.fields[0].get(),
                'codigo': worker_field.fields[1].get(),
                'nombre': worker_field.fields[2].get()
            })

        # Setup initial row with months and days
        header = ['Turno', 'Codigo', 'Nombre']
        days = ['Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sa', 'Do']
        initial_day = days.index(self.day[:2])
        
        months = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        month_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        
        # Generate all day headers
        current_day = 1
        current_month = 0
        day_count = sum(month_days)
        
        for _ in range(day_count):
            header.append(f"{days[(initial_day + current_day - 1) % 7]} {current_day} {months[current_month]}")
            current_day += 1
            if current_day > month_days[current_month]:
                current_day = 1
                current_month += 1
        
        ws.append(header)

        # Generate schedules for first block (workers 0-4)
        for worker_idx in range(5):
            worker = workers[worker_idx]
            row = [worker['turno'], worker['codigo'], worker['nombre']]
            
            current_day = 1
            current_week = 1
            
            for day in range(day_count):
                weekday = (initial_day + day) % 7
                is_weekend = weekday >= 5
                
                # Determine if this worker has N or A this week
                week_position = (current_week + worker_idx) % 5
                has_special = week_position in [0, 1]
                shift = 'TC' if is_weekend and has_special else \
                       'L' if is_weekend else \
                       'N' if has_special and week_position == 0 else \
                       'A' if has_special and week_position == 1 else \
                       'D'
                
                row.append(shift)
                
                if weekday == 0:
                    current_week += 1
            
            ws.append(row)

        # Add empty row between blocks
        ws.append([''] * len(header))

        # Generate schedules for second block (workers 5-8)
        for worker_idx in range(5, 9):
            worker = workers[worker_idx]
            row = [worker['turno'], worker['codigo'], worker['nombre']]
            
            current_day = 1
            current_week = 1
            
            for day in range(day_count):
                weekday = (initial_day + day) % 7
                is_weekend = weekday >= 5
                
                worker_pair = (worker_idx - 5) % 2  # 0 or 1
                is_odd_week = current_week % 2 == 1
                
                if is_weekend:
                    if worker_pair == 0:  # Workers 1&4
                        shift = 'TC' if is_odd_week else 'L'
                    else:  # Workers 2&3
                        shift = 'L' if is_odd_week else 'TC'
                else:
                    shift = 'D'  # Monday to Friday
                
                row.append(shift)
                
                if weekday == 0:
                    current_week += 1
            
            ws.append(row)

        # Save the workbook
        wb.save('cronograma.xlsx')
        print("Cronogram generated successfully!")
        subprocess.Popen(['start', 'cronograma.xlsx'], shell=True)
        os.Exit(0)


app = App()
app.mainloop()
