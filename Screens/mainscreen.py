import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import Calendar
from datetime import datetime
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from Main.MainController import MainController
class MainWindow(tk.Tk):
    def __init__(self):
        self.mainController : "MainController"
        self.product_serial_numbers = []
        super().__init__()
        self.title("Main Screen")
        self.geometry("890x600")
        self.bind_all("<MouseWheel>", self.on_mouse_wheel)
        self.working_order_value = None
        self.starting_shift_value = None
        self.product_list = []
        self.schedule_start = None
        self.schedule_end = None

        self.center_window(900, 600)

        # Ana ekranı ikiye bölmek için bir Frame
        left_frame = tk.Frame(self, width=250, height=450, padx=10, pady=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        left_frame.grid_columnconfigure(0, weight=0)  # Etiket sütunu
        left_frame.grid_columnconfigure(1, weight=1)  # Giriş alanları sütunu
        left_frame.grid_columnconfigure(2, weight=2)
        right_frame = tk.Frame(self, width=400, height=450, padx=10, pady=10)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Sol taraftaki bileşenler

        # Ürün ekleme
        tk.Label(left_frame, text="Product Serial Number:").grid(row=2, column=0, sticky="w", pady=20, padx=5)
        self.product_entry = ttk.Entry(left_frame, width=30)
        self.product_entry.grid(row=2, column=1, sticky="w", pady=5, padx=2)
        self.product_entry.insert(0, "")

        self.add_product_button = ttk.Button(left_frame, width=8, text="Add", command=self.add_product)
        self.add_product_button.grid(row=2, column=3, sticky="w", padx=0, pady=5)

        # Ürün jigi seçme
        self.jig_list = ttk.Combobox(left_frame, width=10)
        self.jig_list.grid(row=2, column=2, sticky="w", pady=10, padx=5)
        self.jig_list.bind("<<ComboboxSelected>>", lambda event: self.jig_list_select())

        self.product_entry.bind("<Return>", lambda event: self.add_product())

        # Select working order
        tk.Label(left_frame, text="Working Order:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.working_order = ttk.Combobox(left_frame, width=27, values=["V1", "V2", "V3"])
        self.working_order.grid(row=1, column=1, sticky="w", pady=10, padx=5)
        self.working_order.bind("<<ComboboxSelected>>", lambda event: self.working_order_select())

        # Excel ekleme butonu
        tk.Button(left_frame, text="Upload Excel",width=17, command=self.upload_excel).grid(row=0, column=0, sticky="w", padx=5, pady=10)

        self.loaded_excel = tk.Entry(left_frame, width=30, state='readonly')
        self.loaded_excel.grid(row=0, column=1, sticky="w", pady=10, padx=5)

        # Edit Jigs butonu
        self.edit_jigs = ttk.Button(left_frame, width=12, text="Edit Jigs", state='disabled', command=self.edit_jigs_screen)
        self.edit_jigs.grid(row=0, column=2, sticky="w", pady=10, padx=5 )
        # butona command eklenip edit jigs ekranı açılacak

        # Takvim ve tarih seçimi
        tk.Label(left_frame, text="Schedule Start Date:").grid(row=4, column=0, sticky="w", pady=10, padx=5)
        self.start_date_entry = tk.Entry(left_frame, width=30)
        self.start_date_entry.grid(row=4, column=1, sticky="w", pady=15, padx=5)

        tk.Label(left_frame, text="Schedule End Date:").grid(row=5, column=0, sticky="w", pady=10, padx=5)
        self.end_date_entry = tk.Entry(left_frame,width=30)
        self.end_date_entry.grid(row=5, column=1, sticky="w", pady=15, padx=5)

        tk.Label(left_frame, text="Start Shift:").grid(row=6, column=0, sticky="w", pady=5, padx=5)
        self.starting_shift = ttk.Combobox(left_frame, width=27, values=["I1", "I2", "I3"])
        self.starting_shift.grid(row=6, column=1, sticky="w", pady=10, padx=5)
        self.starting_shift.bind("<<ComboboxSelected>>", lambda event: self.starting_shift_select())

        self.calendar = Calendar(left_frame)
        self.calendar.grid(row=7, column=0, columnspan=2, pady=15)
        self.calendar.bind("<<CalendarSelected>>", self.select_date)

        # Select function
        tk.Label(left_frame, text="Select Function:").grid(column=0, row=8, sticky="w", pady=15)
        self.function_select = ttk.Combobox(left_frame, width=27,
                                            values=["Assignment with Current Schedule", "Assignment with Overtime"])
        self.function_select.grid(column=1, row=8, sticky="w", pady=15, padx=5)

        # Make assignment button
        tk.Button(left_frame, text="Make Assignment", command=self.make_assignment).grid(column=0, row=9, columnspan=3,pady=1)

        # Sağ taraftaki bileşenler
        p_frame = tk.Frame(right_frame)
        p_frame.pack(fill="both", expand=False, pady=18)

        #başlık
        tk.Label(p_frame, text="Product List:").pack(pady=2, padx=15)

        # canvas ve List scroll
        self.p_canvas = tk.Canvas(p_frame, width=250, height=200, bg="white")
        self.p_canvas.pack(side="left", fill="both", expand=True, pady=5)
        self.p_scrollbar = ttk.Scrollbar(p_frame, orient="vertical", command=self.p_canvas.yview)
        self.p_scrollbar_frame = tk.Frame(self.p_canvas, bg="white")

        self.p_scrollbar_frame.bind("<Configure>", lambda e: self.p_canvas.configure(scrollregion=self.p_canvas.bbox("all")))
        self.p_canvas.create_window((0, 0), window=self.p_scrollbar_frame, anchor="nw")
        self.p_canvas.configure(yscrollcommand=self.p_scrollbar.set)

        self.p_canvas.pack(fill="both", expand=True)
        self.p_scrollbar.pack(side="right", fill="y")

        self.p_canvas.bind("<MouseWheel>", self.on_mouse_wheel)

        # Add Offdays butonu
        tk.Button(right_frame, text="Add Off Days", command=lambda : self.off_days_screen()).pack(pady=5, padx=10)

        # Offday workers list
        w_frame = tk.Frame(right_frame)
        w_frame.pack(fill="both", expand=False)

        self.w_canvas = tk.Canvas(w_frame, width=250, height=200, bg="white")
        self.w_scrollbar = ttk.Scrollbar(w_frame, orient="vertical", command=self.w_canvas.yview)
        self.w_scrollbar_frame = tk.Frame(self.w_canvas, bg="white")

        self.w_scrollbar_frame.bind("<Configure>", lambda e: self.w_canvas.configure(scrollregion=self.w_canvas.bbox("all")))
        self.w_canvas.create_window((0, 0), window=self.w_scrollbar_frame, anchor="nw")
        self.w_canvas.configure(yscrollcommand=self.w_scrollbar.set)

        self.w_canvas.pack(side="left", fill="both", expand=True)
        self.w_scrollbar.pack(side="right", fill="y")

        self.w_canvas.bind("<MouseWheel>", self.on_mouse_wheel)

    def toggle_calendar(self, target):
        print(target)

    def center_window(self, width, height):
        # Ekranın boyutunu al
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Pencereyi ekranın ortasına yerleştir
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def edit_jigs_screen(self):

        edit_jigs_window = tk.Toplevel(self)
        edit_jigs_window.title(f"Edit Jigs")
        edit_jigs_window.geometry("250x500")

        # Canvas and scrollbar
        ej_frame = tk.Frame(edit_jigs_window)
        ej_frame.pack(fill=tk.BOTH, expand=True)
        ej_canvas = tk.Canvas(ej_frame)
        ej_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ej_scrollbar = tk.Scrollbar(ej_frame, orient=tk.VERTICAL, command=ej_canvas.yview)
        ej_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        ej_canvas.configure(yscrollcommand=ej_scrollbar.set)  # Doğru değişken kullanıldı
        ej_canvas.bind("<Configure>", lambda e: ej_canvas.configure(scrollregion=ej_canvas.bbox("all")))

        def mouse_wheel(event):
            ej_canvas.yview_scroll(-1 * int(event.delta / 120), "units")

        ej_canvas.bind_all("<MouseWheel>", mouse_wheel)

        # Frame inside the canvas
        content_frame = tk.Frame(ej_canvas)
        ej_canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Headers
        headers = ["Jigs", "Turn Off Jig"]
        for col, header in enumerate(headers):
            tk.Label(content_frame, text=header, font=("Arial", 12, "bold")).grid(row=0, column=col, padx=30, pady=5)

        checkboxes = []

        jig_len = len(self.mainController.get_jigs())
        for i in range(jig_len):
            jig = self.mainController.get_jigs()[i]
            tk.Label(content_frame, text=f"Jig {jig.get_name()}").grid(row=i + 1, column=0, padx=20, pady=5)

            def on_checkbox_toggle(index=i):
                jig = self.mainController.get_jigs()[index]
                if jig.get_state() == None:
                    jig.set_state(True)
                elif jig.get_state() == True:
                    jig.set_state(None)

            checkbox_var = tk.BooleanVar(value=bool(jig.get_state()))
            tk.Checkbutton(content_frame, variable=checkbox_var, command=on_checkbox_toggle).grid(row=i + 1, column=1,
                                                                                                  padx=5, pady=5)
            checkboxes.append([checkbox_var, jig.get_name()])

        # Save button
        ej_save_button = tk.Button(edit_jigs_window, text="Save Changes", command=lambda : self.save_and_close(edit_jigs_window))
        ej_save_button.pack(side=tk.BOTTOM, pady=10)

        edit_jigs_window.mainloop()

    def save_and_close(self, edit_jigs_window):
        edit_jigs_window.destroy()  # Pencereyi kapat



    def update_jig_list(self):
        current_values = list(self.jig_list['values'])  # Mevcut değerleri al

        for jig in self.mainController.get_jigs():
            current_values.append(jig.get_name())  # Yeni değeri ekle

        self.jig_list.configure(values=current_values)  # Combobox'u güncelle

    def jig_list_select(self):
        selected_jig_name = self.jig_list.get().strip()
        for jig in self.mainController.get_jigs():
            if jig.get_name() == selected_jig_name:
                self.selected_jig = jig

    def working_order_select(self):
        self.working_order_value = self.working_order.get()
        self.mainController.get_data_loader_object().set_working_order(self.working_order_value)

    def starting_shift_select(self):
        self.starting_shift_value = self.starting_shift.get()
        self.mainController.get_data_loader_object().set_starting_shift(self.starting_shift_value)


    def add_product(self):
        serial_number = self.product_entry.get().strip()
        if self.selected_jig.get_state() == None:
            if serial_number and serial_number not in self.product_serial_numbers:
                self.product_serial_numbers.append(serial_number)
                self.update_product_list()
                self.mainController.create_product(serial_number)
                self.mainController.get_product(serial_number).set_current_jig(self.selected_jig)
                self.selected_jig.set_state(True)
                self.selected_jig.set_assigned_product(self.mainController.get_product(serial_number))
                self.product_entry.delete(0, tk.END)

                result =  self.mainController.get_data_loader_object().read_operations_from_excel(serial_number)

                if result != 1:
                    messagebox.showerror("Error", f"Failed to load Excel: {str(result)}")
                else:
                    self.mainController.set_all_previous_operations(serial_number)
                    self.mainController.calculate_required_worker(serial_number)
                    self.mainController.calculate_operating_duration(serial_number)
        else:
            messagebox.showerror("Error", f"Selected {str(self.selected_jig.get_name())} is busy")

    def update_product_list(self):
        # Mevcut ürün listesi widgetlarını temizle
        for widget in self.p_scrollbar_frame.winfo_children():
            widget.destroy()

        # Ürünleri yeniden listele
        for serial_number in self.product_serial_numbers:
            row_frame = ttk.Frame(self.p_scrollbar_frame)
            row_frame.pack(fill="x", pady=2)

            # Delete butonu
            delete_button = tk.Button(row_frame, text="❌",
                                      command=lambda sn=serial_number: self.delete_product(sn),
                                      bd=0, bg="white", fg="red",
                                      activebackground="white", activeforeground="darkred",
                                      font=("Arial", 10, "bold"))
            delete_button.grid(row=0, column=0, padx=5, pady=2)

            # Ürün adı
            product_label = ttk.Label(row_frame, text=serial_number, width=30, anchor="w", background="white")
            product_label.grid(row=0, column=1, padx=5, pady=2)

            # Add Progress butonu
            progress_button = ttk.Button(row_frame, text="Add Progress", command=lambda sn=serial_number: self.progress_update_screen(sn))
            progress_button.grid(row=0, column=2, padx=5, pady=2)

    def delete_product(self, serial_number):
        if serial_number in self.product_serial_numbers:
            self.product_serial_numbers.remove(serial_number)
            self.update_product_list()
            self.mainController.delete_product(serial_number)

    def progress_update_screen(self, serial_number):
        # Progress window
        progress_window = tk.Toplevel(self)
        progress_window.title(f"Progress Update - {serial_number}")
        progress_window.geometry("500x600")

        # Canvas and scrollbar
        pw_frame = tk.Frame(progress_window)
        pw_frame.pack(fill=tk.BOTH, expand=True)
        pw_canvas = tk.Canvas(pw_frame)
        pw_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pw_scrollbar = tk.Scrollbar(pw_frame, orient=tk.VERTICAL, command=pw_canvas.yview)
        pw_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        pw_canvas.configure(yscrollcommand=pw_scrollbar.set)
        pw_canvas.bind("<Configure>", lambda e: pw_canvas.configure(scrollregion=pw_canvas.bbox("all")))

        def mouse_wheel(event):
            pw_canvas.yview_scroll(-1 * int(event.delta / 120), "units")

        pw_canvas.bind_all("<MouseWheel>", mouse_wheel)

        # Frame inside the canvas
        content_frame = tk.Frame(pw_canvas)
        pw_canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Headers
        headers = ["Operation", "Status", "Progress %", "Duration (days)", "Remaining"]
        for col, header in enumerate(headers):
            tk.Label(content_frame, text=header, font=("Arial", 12, "bold")).grid(row=0, column=col, padx=5, pady=5)

        # Progress entries and checkboxes
        checkboxes = []
        progress_entries = []

        product = self.mainController.get_product(serial_number)
        operations = product.get_operations()
        op_len = len(operations)

        for i in range(op_len):
            op = self.mainController.get_product(serial_number).get_operations()[i]

            # Operation name
            tk.Label(content_frame, text=f"Operation {op.get_name()}").grid(row=i + 1, column=0, padx=5, pady=5)

            # Status checkbox
            def on_checkbox_toggle(index=i):
                op = self.mainController.get_product(serial_number).get_operations()[index]
                current_state = op.get_completed()
                new_state = not current_state

                # Checkbox değişimini işle
                if new_state:  # Tamamlandı olarak işaretlendi
                    op.set_completed(new_state)
                    op.set_remaining_duration(0.0)  # Kalan süreyi sıfırla
                    checkboxes[index][0].set(new_state)
                    progress_entries[index][0].delete(0, tk.END)
                    progress_entries[index][0].insert(0, "100")
                    # Kalan süre etiketini güncelle
                    remaining_label = progress_entries[index][1]
                    remaining_label.config(text="0.00")
                    select_predecessors(index)
                else:  # Tamamlanmadı olarak işaretlendi
                    op.set_completed(new_state)
                    op.set_remaining_duration(op.get_operating_duration())  # Kalan süreyi tam süreye eşitle
                    checkboxes[index][0].set(new_state)
                    progress_entries[index][0].delete(0, tk.END)
                    progress_entries[index][0].insert(0, "0")
                    # Kalan süre etiketini güncelle
                    remaining_label = progress_entries[index][1]
                    remaining_label.config(text=f"{op.get_remaining_duration():.2f}")

            checkbox_var = tk.BooleanVar(value=bool(op.get_completed()))
            tk.Checkbutton(content_frame, variable=checkbox_var, command=on_checkbox_toggle).grid(row=i + 1, column=1,
                                                                                                  padx=5, pady=5)
            checkboxes.append([checkbox_var, op.get_name()])

            # Progress percentage entry
            progress_entry = tk.Entry(content_frame, width=10)
            progress_entry.grid(row=i + 1, column=2, padx=5, pady=5)
            if op.get_completed():
                progress_entry.insert(0, "100")
            else:
                # Eğer kalan süre tam süreye eşit değilse, ilerleme hesapla
                total_duration = op.get_operating_duration()
                remaining = op.get_remaining_duration()
                if total_duration > 0 and remaining < total_duration:
                    progress = int((1 - remaining / total_duration) * 100)
                    progress_entry.insert(0, str(progress))
                else:
                    progress_entry.insert(0, "0")

            # Duration (days) display
            duration_days = op.get_operating_duration()
            tk.Label(content_frame, text=f"{duration_days:.2f}").grid(row=i + 1, column=3, padx=5, pady=5)

            # Remaining duration display
            remaining_label = tk.Label(content_frame, text=f"{op.get_remaining_duration():.2f}")
            remaining_label.grid(row=i + 1, column=4, padx=5, pady=5)

            # Store entry and label widgets
            progress_entries.append([progress_entry, remaining_label])

            # Progress değişimini işleme fonksiyonu
            def update_progress(event, index=i):
                entry = progress_entries[index][0]
                try:
                    progress_value = float(entry.get().strip())
                    if progress_value < 0:
                        progress_value = 0
                    elif progress_value > 100:
                        progress_value = 100

                    op = self.mainController.get_product(serial_number).get_operations()[index]
                    # Kalan süreyi hesapla ve güncelle
                    self.mainController.update_operation_remaining_duration(op, progress_value)

                    # Checkbox durumunu güncelle
                    if progress_value == 100:
                        checkboxes[index][0].set(True)
                        op.set_completed(True)
                        # Öncülleri otomatik tamamla
                        select_predecessors(index)
                    else:
                        checkboxes[index][0].set(False)
                        op.set_completed(False)

                    # Kalan süre etiketini güncelle
                    remaining_label = progress_entries[index][1]
                    remaining_label.config(text=f"{op.get_remaining_duration():.2f}")

                except ValueError:
                    # Geçersiz değer girilirse, sıfırla
                    entry.delete(0, tk.END)
                    entry.insert(0, "0")

            # Progress entry'ye değişiklik algılama ekle
            progress_entry.bind("<FocusOut>", lambda event, idx=i: update_progress(event, idx))
            progress_entry.bind("<Return>", lambda event, idx=i: update_progress(event, idx))

        def select_predecessors(index):
            selected_op_name = checkboxes[index][1]
            selected_op = product.get_operation_by_name(selected_op_name)
            selected_op_predecessors = selected_op.get_previous_operations()
            for pre in selected_op_predecessors:
                if not pre.get_completed():  # Eğer öncül tamamlanmamışsa
                    pre.set_completed(True)  # Öncül operasyonu tamamlandı olarak işaretle
                    pre.set_remaining_duration(0)  # Kalan süreyi 0 yap

                    # Öncül operasyonun checkbox'ını güncelle
                    for i, checkbox in enumerate(checkboxes):
                        if checkbox[1] == pre.get_name():
                            checkbox[0].set(True)  # Checkbox'ı işaretle
                            if i < len(progress_entries):
                                progress_entries[i][0].delete(0, tk.END)
                                progress_entries[i][0].insert(0, "100")
                                progress_entries[i][1].config(text="0.00")

        # Save button
        save_button = tk.Button(progress_window, text="Save Changes",
                                command=lambda: self.save_progress_changes(progress_window, serial_number))
        save_button.pack(side=tk.BOTTOM, pady=10)

        progress_window.mainloop()

    def save_progress_changes(self, progress_window, serial_number):
        self.mainController.calculate_product_progress(serial_number)
        self.mainController.remove_completed_predecessors(serial_number)
        self.mainController.set_critical_operations(serial_number)
        progress_window.destroy()

    def saveandclose(self, progress_window):
        progress_window.destroy()  # Pencereyi kapat


    def off_days_screen(self):
        off_day_window = tk.Toplevel(self)
        off_day_window.title("Add Off-days")
        off_day_window.geometry("600x600")
        # Pencere kapatıldığında on_close fonksiyonunu çağır
        off_day_window.protocol("WM_DELETE_WINDOW", lambda: self.close_add_off_days(off_day_window))

        # Main frame for canvas and scrollbar
        main_frame = tk.Frame(off_day_window)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def _on_mouse_wheel(event):
            canvas.yview_scroll(-1 * int(event.delta / 120), "units")

        canvas.bind("<MouseWheel>", _on_mouse_wheel)

        content_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # Header
        tk.Label(content_frame, text="Registration No", font=("Arial", 12, "bold"), width=15).grid(row=0, sticky="w", column=0,
                                                                                                   padx=5, pady=5)
        tk.Label(content_frame, text="Name", font=("Arial", 12, "bold"), width=20).grid(row=0, column=1, sticky="w", padx=5,
                                                                                        pady=5)
        workers = self.mainController.get_workers()
        worker_labels = []

        row_num = 2
        for worker in self.mainController.get_workers():
            # Worker bilgileri
            tk.Label(content_frame, text=worker.get_registration_number(), width=15).grid(row=row_num, column=0, padx=5, pady=5)
            tk.Label(content_frame, text=worker.get_name(), width=20).grid(row=row_num, column=1, padx=5, pady=5)

            # Takvim Butonu (Entry'nin yanına)
            ttk.Button(content_frame, text="📅", command=lambda wrkr=worker: self.open_off_days_date(wrkr),width=3).grid(row=row_num, column=2, padx=5, pady=5)

            row_num += 1

        # Search boxes
        reg_search_entry = ttk.Entry(content_frame, width=30)
        reg_search_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        name_search_entry = ttk.Entry(content_frame, width=30)
        name_search_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Todo filtreleme fonksiyonu eklenecek
        '''
        def filter_workers(event=None):
            search_reg = reg_search_entry.get().strip().lower()
            search_name = name_search_entry.get().strip().lower()

            for widget in content_frame.winfo_children():
                if int(widget.grid_info().get("row", 0)) > 1:
                    widget.destroy()'''

        off_day_window.mainloop()

    def close_add_off_days(self, window):
        window.destroy()  # Pencereyi kapat
        # Mevcut ürün listesi widgetlarını temizle
        for widget in self.w_scrollbar_frame.winfo_children():
            widget.destroy()
        self.update_off_days_list()



    def update_off_days_list(self):
        for worker in self.mainController.get_workers():
            if worker.get_off_days() != None:
                # Ürünleri yeniden listele
                row_frame = ttk.Frame(self.w_scrollbar_frame)
                row_frame.pack(fill="x", pady=2)

                # Delete butonu
                delete_button = tk.Button(row_frame, text="❌",
                                          command=lambda wrkr=worker: self.delete_worker_off_day(wrkr),
                                          bd=0, bg="white", fg="red",
                                          activebackground="white", activeforeground="darkred",
                                          font=("Arial", 10, "bold"))
                delete_button.grid(row=0, column=0, padx=5, pady=2)

                # Worker adı
                worker_name_label = ttk.Label(row_frame, text=worker.get_name(), anchor="w", background="white")
                worker_name_label.grid(row=0, column=1, padx=5, pady=2)

                # Worker off_day_start
                worker_off_day_start = ttk.Label(row_frame, text=worker.get_off_days()[0], width=10, anchor="w", background="white")
                worker_off_day_start.grid(row=0, column=2, padx=5, pady=2)

                # Worker off_days_end
                worker_off_day_end = ttk.Label(row_frame, text=worker.get_off_days()[1], width=10, anchor="w", background="white")
                worker_off_day_end.grid(row=0, column=3, padx=5, pady=2)

    def delete_worker_off_day(self, worker):
        for widget in self.w_scrollbar_frame.winfo_children():
            widget.destroy()
        worker.set_off_days(None)
        self.update_off_days_list()

    def open_off_days_date(self, worker):
        self.open_off_days_date_screen = tk.Toplevel(self)
        self.open_off_days_date_screen.title("Date Frame")
        self.open_off_days_date_screen.geometry("600x400")

        # Main frame for canvas and scrollbar
        main_frame = tk.Frame(self.open_off_days_date_screen)
        main_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Takvim ve tarih seçimi
        tk.Label(canvas, text="Start Date:").grid(row=0, column=0, sticky="w", pady=10, padx=5)
        self.off_start_date_entry = tk.Entry(canvas, width=30)
        self.off_start_date_entry.grid(row=0, column=1, sticky="w", pady=10, padx=5)

        tk.Label(canvas, text="End Date:").grid(row=0, column=2, sticky="w", pady=10, padx=5)
        self.off_end_date_entry = tk.Entry(canvas,width=30)
        self.off_end_date_entry.grid(row=0, column=3, sticky="w", pady=10, padx=5)

        self.offDays = []

        self.calendar = Calendar(canvas)
        self.calendar.grid(row=1, column=0, columnspan=4, pady=15)
        self.calendar.bind("<<CalendarSelected>>", self.select_off_days_date)

        # Save button (Test amaçlı)
        save_button = tk.Button(canvas, text="Save Changes", command=lambda wrkr=worker: self.save_off_days(worker))
        save_button.grid(row=2, column=0, columnspan=4, pady=15)

        self.open_off_days_date_screen.mainloop()

    def select_off_days_date(self, event):
        selected_date = self.calendar.get_date()  # Seçilen tarihi al
        formatted_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%d.%m.%Y")
        if not self.off_start_date_entry.get():
            self.off_start_date_entry.insert(0, formatted_date)
        else:
            self.off_end_date_entry.delete(0, tk.END)  # Önceki veriyi temizle
            self.off_end_date_entry.insert(0, formatted_date)
        self.offDays = [self.off_start_date_entry.get(), self.off_end_date_entry.get()]

    def save_off_days(self, worker):
        worker.set_off_days(self.offDays)
        print( str(worker.get_name()) + str(worker.get_off_days()) )
        self.open_off_days_date_screen.destroy()


    def upload_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select a File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            # Sadece dosya adını uzantısız göster
            file_name = file_path.split('/')[-1].split('.')[0]
            display_name = file_name if len(file_name) <= 30 else '...' + file_name[-27:]

            self.loaded_excel.config(state='normal')  # Giriş alanını düzenlenebilir yap
            self.loaded_excel.delete(0, tk.END)  # Önceki içeriği temizle
            self.loaded_excel.insert(0, display_name)  # Dosya adını göster
            self.loaded_excel.config(state='readonly')  # Yalnızca okunabilir yap

            # Edit Jigs butonunu aktif hale getir
            self.edit_jigs.config(state='normal')


        if file_path:
            result = self.mainController.get_data_loader_object().read_jigs_from_excel(file_path)
            if result == 1:
                self.update_jig_list()
            else:
                messagebox.showerror("Error", f"Failed to load Excel: {str(result)}")

        if file_path:
            result = self.mainController.get_data_loader_object().read_workers_from_excel(file_path)
            if not result == 1:
                messagebox.showerror("Error", f"Failed to load Excel: {str(result)}")

        self.mainController.debug()


    def select_date(self, event):
        selected_date = self.calendar.get_date()  # Seçilen tarihi al
        formatted_date = datetime.strptime(selected_date, "%m/%d/%y")

        if not self.start_date_entry.get():
            self.start_date_entry.insert(0, formatted_date.strftime("%d.%m.%Y"))
            self.schedule_start = formatted_date
        else:
            self.end_date_entry.delete(0, tk.END)  # Önceki veriyi temizle
            self.end_date_entry.insert(0, formatted_date.strftime("%d.%m.%Y"))
            self.schedule_end = formatted_date

    def make_assignment(self):
        print("Making Assignment")
        self.mainController.set_schedule_attributes()
        self.mainController.make_assignment_preparetions()

        # Çıktı verme kısmısı
        assignments = self.mainController.get_assignments_for_output()

        if assignments:
            # Atamalar başarılıysa atama sonuç penceresini göster
            assignment_window = AssignmentOutputWindow(assignments)

            # Excel export butonu ekle
            export_button = tk.Button(
                assignment_window.root,
                text="Export to Excel",
                command=lambda: self.export_to_excel(assignment_window)
            )
            export_button.pack(side=tk.BOTTOM, pady=5)

            # Gantt Chart butonu ekle
            gantt_button = tk.Button(
                assignment_window.root,
                text="Export Gantt Chart",
                command=lambda: self.export_gantt_to_excel(assignment_window)
            )
            gantt_button.pack(side=tk.BOTTOM, pady=5)

    def export_to_excel(self, assignment_window=None):
        """
        Atama sonuçlarını Excel'e aktarır
        """
        success = self.mainController.export_assignments_to_excel()

        if success:
            messagebox.showinfo("Export Successful", "Assignments successfully exported to Excel file.")
        else:
            messagebox.showerror("Export Failed", "Failed to export assignments to Excel.")

    def export_gantt_to_excel(self, assignment_window=None):
        """
        Sadece Gantt şemasını ayrı bir Excel dosyasına aktarır
        """
        success = self.mainController.export_gantt_chart_to_excel()

        if success:
            messagebox.showinfo("Export Successful", "Gantt chart successfully exported to Excel file.")
        else:
            messagebox.showerror("Export Failed", "Failed to export Gantt chart to Excel.")
    def calculate_prdct_prgrss(self, serial_number):
        self.mainController.calculate_product_progress(serial_number)

    def on_mouse_wheel(self, event):
        if self.p_canvas.winfo_ismapped():
            self.p_canvas.yview_scroll(-1 * (event.delta // 120), "units")
        if self.w_canvas.winfo_ismapped():
            self.w_canvas.yview_scroll(-1 * (event.delta // 120), "units")

    def get_schedule_end(self):
        return self.schedule_end

    def get_schedule_start(self):
        return self.schedule_start

    def get_starting_shift(self):
        return self.starting_shift_value

    def get_working_order_value(self):
        return self.working_order_value


    def setMainController(self, mainControllerObject):
        self.mainController = mainControllerObject

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()


class AssignmentOutputWindow:
    def __init__(self, assignments):
        self.assignments = assignments
        self.root = tk.Tk()
        self.root.title("Atama Çıktısı")
        self.root.geometry("1000x600")

        # Treeview widget'ını oluştur
        self.tree = ttk.Treeview(self.root,
                                 columns=("Product", "Jig", "Operation", "Date", "Shift", "Time Interval", "Workers"),
                                 show="headings")
        self.tree.heading("Product", text="Product")
        self.tree.heading("Jig", text="Jig")
        self.tree.heading("Operation", text="Operation")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Shift", text="Shift")
        self.tree.heading("Time Interval", text="Time Interval")
        self.tree.heading("Workers", text="Workers")
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Scrollbar ekle
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Button Frame oluştur
        button_frame = tk.Frame(self.root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        # Export butonları için boşluk bırak - mainscreen.py'deki butonlar buraya eklenecek

        # Atamaları Treeview'a ekle
        self.populate_treeview()

    def populate_treeview(self):
        # Aynı product'a ait atamaları grupla
        product_assignments = {}
        for assignment in self.assignments:
            product_serial = assignment["Product"]
            if product_serial not in product_assignments:
                product_assignments[product_serial] = []
            product_assignments[product_serial].append(assignment)

        # Treeview'a atamaları ekle
        for product_serial, assignments in product_assignments.items():
            # Operasyon adına göre sırala (operasyon adları genellikle sayısal olduğu için)
            sorted_assignments = sorted(assignments,
                                        key=lambda x: int(x["Operation"]) if x["Operation"].isdigit() else x[
                                            "Operation"])

            for assignment in sorted_assignments:
                self.tree.insert("", tk.END, values=(
                    assignment["Product"],
                    assignment["Jig"],
                    assignment["Operation"],
                    assignment["Date"],
                    assignment["Shift"],
                    assignment["Time Interval"],
                    assignment["Workers"]
                ))