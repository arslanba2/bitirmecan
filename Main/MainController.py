from Screens import mainscreen
from Models.Dictionaries import SKILLS
from Models import Product
from Functions import SetCriticalOperation, WorkerAssigner
from Functions.ExcelDataLoader import ExcelDataLoader
from datetime import datetime


class MainController:
    def __init__(self):
        self.__screenController = None
        self.__products = []  # Holds product objects
        self.__jigs = []  # Holds jig objects
        self.__workers = []  # Holds worker objects
        self.__all_critical_operations = []  # [(product , crital operations)]
        self.__dataLoaderObject = ExcelDataLoader()  # Creates DataLoaderObject
        self.__dataLoaderObject.set_products(self.__products)
        self.__dataLoaderObject.set_jigs(self.__jigs)
        self.__dataLoaderObject.set_workers(self.__workers)
        self.__ScheduleObject = WorkerAssigner.Schedule()
        self.__critical_op_check_list = []

    def create_product(self, serialNumber=None):
        Product.create_product(self.__products, serialNumber)

    def delete_product(self, serialNumber):
        for product in self.__products:
            if product.get_serial_number() == serialNumber:
                product.get_current_jig().set_state(None)
                product.get_current_jig().set_assigned_product(None)
                self.__products.remove(product)

    def get_product_list(self):
        return self.__products

    def get_product(self, serialNumber):
        for product in self.__products:
            if product.get_serial_number() == serialNumber:
                return product

    def get_jig(self, name):
        for jig in self.__jigs:
            if jig.get_name() == name:
                return jig

    def get_jigs(self):
        return self.__jigs

    def get_jig(self, _name):
        for jig in self.__jigs:
            if jig.get_name() == _name:
                return jig

    def get_workers(self):
        return self.__workers

    def get_worker(self, _reg_no):
        for worker in self.__workers:
            if worker.get_registration_number() == _reg_no:
                return worker

    def get_data_loader_object(self):
        return self.__dataLoaderObject

    def get_ScheduleObject(self):
        return self.__ScheduleObject

    def calculate_required_worker(self, sn):
        product = self.get_product(sn)
        for operation in product.get_operations():
            required_man = operation.get_min_workers()
            while required_man <= operation.get_max_workers():
                if (operation.get_required_man_hours()/required_man) <= 7.5:
                    operation.set_required_worker(required_man)
                    break
                else:
                    required_man = required_man + 1
            print(f"Op {operation.get_name()} req shift: {operation.get_required_worker()}")

    def calculate_operating_duration(self, sn):
        product = self.get_product(sn)
        for operation in product.get_operations():
            duration = operation.get_required_man_hours()/(7.5*operation.get_required_worker())
            operation.set_operating_duration(duration)

        self.print_operation_durations()

    def print_operation_durations(self):
        for product in self.__products:
            for operation in product.get_operations():
                print(f"Operation {operation.get_name()} duration: {operation.get_operating_duration()}")

    def set_all_previous_operations(self, sn):
        product = self.get_product(sn)
        all_ops = product.get_operations()
        for op in product.get_operations():
            self.find_all_previous_operations(op, all_ops)

    def find_all_previous_operations(self, operation, all_operations, visited=None):
        if visited is None:
            visited = set()

        if operation.get_name() in visited:
            return

        visited.add(operation.get_name())
        prev_list = operation.get_previous_operations()

        for pred in operation.get_predecessors():
            if pred not in prev_list:
                prev_list.append(pred)
            self.find_all_previous_operations(pred, all_operations, visited)
            for pred_prev in pred.get_previous_operations():
                if pred_prev not in prev_list:
                    prev_list.append(pred_prev)

        operation.set_previous_operations(prev_list)

    def calculate_product_progress(self, serial_number):
        product = self.get_product(serial_number)
        total_duration = 0
        for operation in product.get_operations():
            total_duration = total_duration + operation.get_required_man_hours()
        applied_duration = 0
        for operation in product.get_operations():
            if operation.get_completed():
                applied_duration = applied_duration + operation.get_required_man_hours()
        progress = applied_duration/total_duration*100
        product.set_progress(progress)

        print(f"product {product.get_serial_number()} progress % : {product.get_progress()}")

    def remove_completed_predecessors(self, _sn):
        product = self.get_product(_sn)
        for operation in product.get_operations():
            if operation.get_completed():  # Tamamlanmış operasyon ise
                successor_of_completed_operation = operation.get_successors()  # Tamamlanmış operasyonun ardıl listesi (str)
                opstoremove = operation.get_previous_operations()
                opstoremove.append(operation)
                for suc_op in successor_of_completed_operation:  # Her bir ardıl operasyon objesini dön
                    suc_op_object = product.get_operation_by_name(suc_op)
                    # Öncül listesini kopyala ve tamamlanmış öncülü çıkar
                    uncomplete_predecessors_names = []
                    for pre in suc_op_object.get_predecessors():
                        uncomplete_predecessors_names.append(pre)
                    for removeop in opstoremove:
                        if removeop in uncomplete_predecessors_names:
                            uncomplete_predecessors_names.remove(removeop)
                    suc_op_object.set_uncompleted_predecessors(uncomplete_predecessors_names)

    # CPM calculation
    def set_critical_operations(self, _sn):
        tasks = []
        product = self.get_product(_sn)
        calculator = SetCriticalOperation.Graph()
        for operation in product.get_operations():
            if not operation.get_completed():
                task = operation.get_name()
                duration = operation.get_operating_duration()
                dependencies = []
                for op in operation.get_uncompleted_predecessors():
                    dependencies.append(op.get_name())
                calculator.add_task(task, duration, dependencies)
                tasks.append(task)

        critical_operations, earliest_start, latest_finish = calculator.find_critical_operations()

        print("Kritik Operasyonlar:", critical_operations)
        print("En Erken Başlama Zamanları:", earliest_start)
        print("En Geç Tamamlanma Zamanları:", latest_finish)

        critical_op_obj_list = []
        for op_name in critical_operations:
            op_obj = product.get_operation_by_name(op_name)
            op_obj.set_early_start(earliest_start[op_name])
            op_obj.set_late_finish(latest_finish[op_name])
            if op_obj.get_early_start() == 0:
                critical_op_obj_list.append(op_obj)
            product.append_critical_operations(critical_op_obj_list)

    def sort_operations_by_duration(self):
        for product in self.__products:
            criticalops = product.get_critical_operations()
            sorted_criticalops = sorted(criticalops, key=lambda op: op.get_operating_duration())
            product.append_critical_operations(sorted_criticalops)

    def sort_products_by_progress(self):
        sorted_products = sorted(self.__products, key=lambda product: product.get_progress() or 0, reverse=True)
        self.__products = sorted_products

    def get_all_critical_operations(self):
        critical_ops = []
        for product in self.__products:
            for op in product.get_critical_operations():
                if len(op.get_uncompleted_predecessors()) == 0:
                    critical_ops.append((product, op))
                    print(f"critical op {op.get_name()}")
        return critical_ops

    def set_schedule_attributes(self):
        self.__ScheduleObject.set_start_date(self.screenController.get_schedule_start())
        self.__ScheduleObject.set_end_date(self.screenController.get_schedule_end())
        self.__ScheduleObject.set_start_shift(self.screenController.get_starting_shift())
        self.__ScheduleObject.set_working_order(self.screenController.get_working_order_value())
        self.__ScheduleObject.create_time_intervals()
        self.assign_workers_to_time_intervals()

    def assign_workers_to_time_intervals(self):
        if not self.__ScheduleObject or not self.__workers:
            raise ValueError("Schedule and workers must be set.")

        for date_obj in self.__ScheduleObject.dates:
            for time_interval in date_obj.time_intervals:
                # TimeInterval'ın tarih, vardiya ve zaman aralığını al
                interval_date = date_obj.get_date()  # Zaman aralığının tarihi
                interval_shift = time_interval.shift  # Vardiya
                interval_start_time = time_interval.interval[0]  # Zaman aralığının başlangıcı
                interval_end_time = time_interval.interval[1]  # Zaman aralığının bitişi

                # Uygun çalışanları bul
                available_workers = []
                for worker in self.__workers:
                    # Çalışanın off-day'lerini kontrol et
                    off_days = worker.get_off_days()
                    if off_days:
                        off_start_date = datetime.strptime(off_days[0], "%d.%m.%Y").date()
                        off_end_date = datetime.strptime(off_days[1], "%d.%m.%Y").date()
                        if off_start_date <= interval_date <= off_end_date:
                            continue  # Çalışan bu tarihte off-day'de, atama yapma

                    for schedule_entry in worker.get_shift_schedule():
                        schedule_date, schedule_shift, available_hours = schedule_entry

                        # Tarih ve vardiya eşleşiyor mu kontrol et
                        if schedule_date == interval_date and schedule_shift == interval_shift:
                            # Zaman aralığı da uyuyor mu kontrol et
                            for hours in available_hours:
                                if hours[0] <= interval_start_time and hours[1] >= interval_end_time:
                                    available_workers.append(worker)
                                    break

                # TimeInterval'ın available_workers listesini güncelle
                time_interval.available_workers = available_workers

    def initiate_assignment(self, critical_op_list):
        # Check if we need to initialize or reset the attempts tracking
        if not hasattr(self, '_assignment_attempts'):
            self._assignment_attempts = {}

        # Update attempt counter for each operation in the critical list
        for product, operation in critical_op_list:
            op_key = (product.get_serial_number(), operation.get_name())
            self._assignment_attempts[op_key] = self._assignment_attempts.get(op_key, 0) + 1

            # If we've tried this operation too many times (e.g., 3 attempts), mark it as completed
            # to prevent it from being selected again, but log this for debugging
            if self._assignment_attempts[op_key] > 3:
                print(f"Warning: Operation {operation.get_name()} for product {product.get_serial_number()} "
                      f"has been attempted {self._assignment_attempts[op_key]} times without success. "
                      f"Marking as completed to prevent infinite loop.")
                operation.set_completed(True)

        # Check if there are any operations left to assign
        unassigned_ops_exist = False
        for product in self.__products:
            for op in product.get_operations():
                if not op.get_completed() and len(op.get_uncompleted_predecessors()) == 0:
                    unassigned_ops_exist = True
                    break
            if unassigned_ops_exist:
                break

        # If no operations can be assigned, we're done
        if not unassigned_ops_exist:
            print("No more operations can be assigned")
            return

        # Check for empty critical list but assignable operations exist
        if not critical_op_list and unassigned_ops_exist:
            print("Critical operation list is empty but assignable operations exist - finding operations to assign")
            assignable_ops = []
            for product in self.__products:
                for op in product.get_operations():
                    if not op.get_completed() and len(op.get_uncompleted_predecessors()) == 0:
                        assignable_ops.append((product, op))

            if assignable_ops:
                critical_op_list = assignable_ops

        # Update the check list for the next iteration
        self.__critical_op_check_list = critical_op_list

        # Original assignment logic continues from here
        op_list = critical_op_list
        for product, operation in op_list:
            # Skip if this operation has been artificially marked as completed
            if operation.get_completed():
                continue

            intervals_list = self.get_ScheduleObject().get_sorted_time_intervals()
            # 1. Önceki operasyonların en geç bitiş zamanını bul
            latest_finish_time = self.find_latest_finish_time_of_predecessors(operation)
            # 2. Interval listesini en geç bitiş zamanından sonraki aralıklarla sınırla
            filtered_intervals = self.filter_intervals_after_time(intervals_list, latest_finish_time)

            # If no suitable intervals found, continue to next operation
            if not filtered_intervals:
                print(
                    f"No suitable intervals found for operation {operation.get_name()} of product {product.get_serial_number()}")
                continue

            assignment_made = False  # Flag to track if we successfully made an assignment

            for interval in filtered_intervals:
                # 1. Önceki operasyonların bu aralıkta olup olmadığını kontrol et
                if not self.previous_operation_control(operation, interval):
                    continue  # Önceki operasyonlar bu aralıkta, atama yapılamaz, sonraki interval'a geç

                # 2. Aralıkta aynı ürüne ait başka bir operasyon olup olmadığını kontrol et
                if self.same_product_control(product, interval):
                    # Aynı ürüne ait başka bir operasyon varsa, jig kapasitesini kontrol et
                    if not self.check_jig_capacity(product, operation, interval):
                        continue  # Jig kapasitesi aşılıyor, atama yapılamaz, sonraki interval'a geç

                    # Jig kapasitesi uygunsa, yeterli sayıda çalışan olup olmadığını kontrol et
                    if not self.compatible_worker_number_check(operation, interval):
                        continue  # Yeterli çalışan yok, atama yapılamaz, sonraki interval'a geç

                    # Tüm kontroller başarılı, atama yap
                    available_workers = interval.get_available_workers()
                    # Worker'ları öncelik sırasına göre sırala
                    sorted_workers = sorted(
                        available_workers,
                        key=lambda worker: self.get_skill_priority(worker.get_skills())
                    )

                    # Check if we have enough workers
                    if len(sorted_workers) < operation.get_required_worker():
                        continue  # Not enough workers, try next interval

                    workers = sorted_workers[:operation.get_required_worker()]
                    jig = product.get_current_jig()

                    # Operasyonun süresi tamamlanana kadar arka arkaya interval'lara atama yap
                    remaining_duration = operation.get_operating_duration()
                    current_interval = interval
                    assignment_intervals = []  # Atama yapılacak interval'ları tutar

                    while remaining_duration > 0:
                        # Ardışık interval'lar için de aynı kontrolleri yap
                        if not self.previous_operation_control(operation, current_interval):
                            break  # Önceki operasyonlar bu aralıkta, atama yapılamaz, döngüden çık

                        if self.same_product_control(product, current_interval):
                            if not self.check_jig_capacity(product, operation, current_interval):
                                break  # Jig kapasitesi aşılıyor, atama yapılamaz, döngüden çık

                            if not self.compatible_worker_number_check(operation, current_interval):
                                break  # Yeterli çalışan yok, atama yapılamaz, döngüden çık

                        # Tüm kontroller başarılı, interval'ı atama listesine ekle
                        assignment_intervals.append(current_interval)
                        remaining_duration -= 0.25  # Her interval 0.25 saat ekler

                        # Bir sonraki interval'ı bul
                        next_interval = self.get_next_interval(current_interval, filtered_intervals)
                        if next_interval is None:
                            # Sonraki interval yok, bu operasyon için atama yapılamaz
                            break  # Döngüden çık ve bir sonraki operasyona geç
                        current_interval = next_interval

                    if remaining_duration <= 0:
                        # Tüm aralıklar uygun, atama yap
                        for assigned_interval in assignment_intervals:
                            self.create_assignment(assigned_interval, jig, product, operation, workers)
                        assignment_made = True  # Successfully made an assignment
                        break  # Operasyonun süresi tamamlandı, bir sonraki operasyona geç

                else:
                    # Aralıkta aynı ürüne ait başka bir operasyon yoksa, yeterli sayıda çalışan kontrolü yap
                    if not self.compatible_worker_number_check(operation, interval):
                        continue  # Yeterli çalışan yok, atama yapılamaz, sonraki interval'a geç

                    # Jig uygunluğunu kontrol et
                    if not self.jig_compatibility_control(product, operation):
                        # Jig uygun değilse, jig değiştir
                        self.change_jig(product, operation)

                    # Tüm kontroller başarılı, atama yap
                    available_workers = interval.get_available_workers()
                    # Worker'ları öncelik sırasına göre sırala
                    sorted_workers = sorted(
                        available_workers,
                        key=lambda worker: self.get_skill_priority(worker.get_skills())
                    )

                    # Check if we have enough workers
                    if len(sorted_workers) < operation.get_required_worker():
                        continue  # Not enough workers, try next interval

                    workers = sorted_workers[:operation.get_required_worker()]
                    jig = product.get_current_jig()

                    # Operasyonun süresi tamamlanana kadar arka arkaya interval'lara atama yap
                    remaining_duration = operation.get_operating_duration()
                    current_interval = interval
                    assignment_intervals = []  # Atama yapılacak interval'ları tutar

                    while remaining_duration > 0:
                        # Ardışık interval'lar için de aynı kontrolleri yap
                        if not self.previous_operation_control(operation, current_interval):
                            break  # Önceki operasyonlar bu aralıkta, atama yapılamaz, döngüden çık

                        if self.same_product_control(product, current_interval):
                            if not self.check_jig_capacity(product, operation, current_interval):
                                break  # Jig kapasitesi aşılıyor, atama yapılamaz, döngüden çık

                            if not self.compatible_worker_number_check(operation, current_interval):
                                break  # Yeterli çalışan yok, atama yapılamaz, döngüden çık

                        # Tüm kontroller başarılı, interval'ı atama listesine ekle
                        assignment_intervals.append(current_interval)
                        remaining_duration -= 0.25  # Her interval 0.25 saat ekler

                        # Bir sonraki interval'ı bul
                        next_interval = self.get_next_interval(current_interval, filtered_intervals)
                        if next_interval is None:
                            # Sonraki interval yok, bu operasyon için atama yapılamaz
                            break  # Döngüden çık ve bir sonraki operasyona geç
                        current_interval = next_interval

                    if remaining_duration <= 0:
                        # Tüm aralıklar uygun, atama yap
                        for assigned_interval in assignment_intervals:
                            self.create_assignment(assigned_interval, jig, product, operation, workers)
                        assignment_made = True  # Successfully made an assignment
                        break  # Operasyonun süresi tamamlandı, bir sonraki operasyona geç

            # If no assignment was made for this operation after trying all intervals,
            # log this fact for debugging
            if not assignment_made:
                print(
                    f"Warning: Could not assign operation {operation.get_name()} for product {product.get_serial_number()}")

        # Continue with the assignment preparation for the next round
        self.make_assignment_preparetions()

    def get_next_interval(self, current_interval, intervals_list):
        """
        Mevcut interval'dan sonraki interval'ı bulur.
        """
        current_index = intervals_list.index(current_interval)
        if current_index + 1 < len(intervals_list):
            return intervals_list[current_index + 1]
        return None  # Sonraki interval yok

    def find_latest_finish_time_of_predecessors(self, operation):
        op = operation
        latest_finish_time = None
        intervals = self.get_ScheduleObject().get_sorted_time_intervals()

        if len(op.get_previous_operations()) == 0:
            first_interval = self.get_ScheduleObject().get_sorted_time_intervals()[0]
            latest_finish_time = (first_interval.get_date(), first_interval.interval[0])
        else:
            for interval in intervals:
                for prev_op in op.get_previous_operations():
                    if prev_op.get_end_datetime():  # Önceki operasyonun bitiş zamanı varsa
                        if latest_finish_time is None or prev_op.get_end_datetime()[0] > latest_finish_time[0]:
                            latest_finish_time = prev_op.get_end_datetime()
                        if latest_finish_time[0] == prev_op.get_end_datetime()[0] and prev_op.get_end_datetime()[1] > \
                                latest_finish_time[1]:
                            latest_finish_time = prev_op.get_end_datetime()
                    else:
                        latest_finish_time = (interval.get_date(), interval.interval[0])

        return latest_finish_time

    def filter_intervals_after_time(self, intervals_list, start_time):
        if start_time is None:
            return intervals_list  # Eğer başlangıç zamanı yoksa, tüm interval listesini döndür
        filtered_intervals = []
        for interval in intervals_list:
            interval_date = interval.get_date()
            interval_start_time = interval.interval[0]  # Interval'ın başlangıç zamanı
            if interval_date > start_time[0]:
                filtered_intervals.append(interval)
            if interval_date == start_time[0] and interval_start_time >= start_time[1]:
                filtered_intervals.append(interval)

        return filtered_intervals

    def get_skill_priority(self, skill):
        priority_order = {
            "KISMİ": 1,
            "KALİTE": 2,
            "ÜRETİM DIŞI": 3,
            "ÜRETİM": 4,
            "HEPSİ": 5
        }
        return priority_order.get(skill, 6)  # Eğer skill tanımlı değilse en düşük öncelik

    def previous_operation_control(self, operation, time_interval):
        op = operation
        interval = time_interval
        assignments = interval.get_assignments()

        # Eğer assignments listesi boşsa, önceki operasyonlar bu aralıkta yok demektir.
        if not assignments:
            return True  # Atama yapılabilir

        # Eğer liste boş değilse, önceki operasyonları kontrol et
        for prev_op in op.get_previous_operations():
            # assignments[2]'nin bir liste veya iterable olup olmadığını kontrol et
            if len(assignments) > 2 and isinstance(assignments[2], (list, tuple)) and prev_op in assignments[2]:
                return False  # Atama yapılamaz

        return True  # Atama yapılabilir

    def same_product_control(self, product, time_interval):  # Buraya operasyonun ait olduğu product gönderilecek
        assignments = time_interval.get_assignments()
        for assignment in assignments:
            if product == assignment[1]:
                return True  # aralıkta aynı product'a ait operasyon var
        return False  # aralıkta aynı product'a ait operasyonn yok

    def jig_compatibility_control(self, product, operation):
        if product.get_current_jig() in operation.get_compatible_jigs():
            return True
        else:
            return False

    def change_jig(self, product, operation):
        jig_names = operation.get_compatible_jigs()
        jig_object_list = []
        for jig_name in jig_names:
            jig_object_list.append(self.get_jig(jig_name))
        for jig in jig_object_list:
            if not jig.get_state():
                jig.set_state(True)
                product.set_current_jig(jig)
            else:
                break

    def check_jig_capacity(self, product, operation, time_interval):

        jig = product.get_current_jig()
        total_workers_assigned = 0

        # Interval içindeki mevcut atamaları kontrol et
        for assignment in time_interval.get_assignments():
            assigned_jig, assigned_product, assigned_operation, assigned_workers = assignment

            # Eğer atama aynı jig ve aynı ürüne aitse, işçi sayısını ekle
            if assigned_jig == jig and assigned_product == product:
                total_workers_assigned += len(assigned_workers)

        # Yeni operasyon için gereken işçi sayısını ekleyerek toplamı kontrol et
        if total_workers_assigned + operation.get_required_worker() <= jig.get_max_assigned_worker():
            return True  # Jig kapasitesi aşılmıyor, atama yapılabilir
        else:
            return False  # Jig kapasitesi aşılıyor, atama yapılamaz

    def compatible_worker_number_check(self, operation, time_interval):
        required_skills = operation.get_required_skills()
        available_workers = []

        for w in time_interval.available_workers:
            worker_skills = w.get_skills()

            # worker_skills'i SKILLS sözlüğünde key olarak ara
            if worker_skills in SKILLS:
                # worker_skills'in karşılığındaki yetenek kümesini al
                worker_skill_set = SKILLS[worker_skills]

                # required_skills, worker_skill_set içinde var mı kontrol et
                if required_skills in worker_skill_set:
                    available_workers.append(w)
            else:
                # worker_skills SKILLS sözlüğünde yoksa, doğrudan eşleşme kontrolü yap
                if required_skills == worker_skills:
                    available_workers.append(w)

        # Yeterli sayıda uyumlu çalışan var mı kontrol et
        if len(available_workers) >= operation.get_required_worker():
            return True  # Atama yapılabilir
        else:
            return False  # Atama yapılamaz

    def create_assignment(self, time_interval, jig, product, operation, workers):
        inter = time_interval
        assignment_entry = jig, product, operation, workers
        inter.set_assignments(assignment_entry)
        jig.set_state(True)
        product.set_current_jig(jig)
        operation.set_completed(True)
        operation.set_start_datetime(inter.get_date(), inter.interval[0])
        operation.set_end_datetime(inter.get_date(), inter.interval[1])
        self.update_worker_shift_schedule(workers, inter)
        self.assign_workers_to_time_intervals()


    def update_worker_shift_schedule(self, workers, time_interval):
        """
        Atama yapılan çalışanların shift schedule'ını günceller.
        Atama yapılan zaman aralığını çalışanların schedule'ından çıkarır.
        """
        interval_start = time_interval.interval[0]  # Zaman aralığının başlangıcı
        interval_end = time_interval.interval[1]  # Zaman aralığının bitişi
        interval_date = time_interval.get_date()  # Zaman aralığının tarihi
        interval_shift = time_interval.shift  # Zaman aralığının vardiyası

        for worker in workers:
            # Çalışanın shift schedule'ını al
            shift_schedule = worker.get_shift_schedule()

            # Shift schedule'ı güncelle
            updated_schedule = []
            for schedule_entry in shift_schedule:
                schedule_date, schedule_shift, available_hours = schedule_entry

                # Eğer tarih ve vardiya eşleşiyorsa, zaman aralığını çıkar
                if schedule_date == interval_date and schedule_shift == interval_shift:
                    new_available_hours = []
                    for hours in available_hours:
                        # Zaman aralığını çıkar
                        if not (hours[0] <= interval_start and hours[1] >= interval_end):
                            new_available_hours.append(hours)

                    # Eğer yeni available_hours boş değilse, schedule'a ekle
                    if new_available_hours:
                        updated_schedule.append((schedule_date, schedule_shift, new_available_hours))
                else:
                    # Tarih ve vardiya eşleşmiyorsa, schedule'ı olduğu gibi ekle
                    updated_schedule.append(schedule_entry)

            # Çalışanın shift schedule'ını güncelle
            worker.set_shift_schedule(updated_schedule)

    def make_assignment_preparetions(self):
        plist = self.__products
        for product in plist:
            sn = product.get_serial_number()
            self.calculate_product_progress(sn)
            self.remove_completed_predecessors(sn)
            self.set_critical_operations(sn)
        self.sort_operations_by_duration()
        self.sort_products_by_progress()
        self.initiate_assignment(self.get_all_critical_operations())

    def get_assignments_for_output(self):
        assignments = []
        for date_obj in self.__ScheduleObject.dates:
            for time_interval in date_obj.time_intervals:
                for assignment in time_interval.get_assignments():
                    jig, product, operation, workers = assignment
                    # Time interval'ları birleştirerek kapsayıcı bir aralık oluştur
                    start_time = time_interval.interval[0].strftime("%H:%M")
                    end_time = time_interval.interval[1].strftime("%H:%M")
                    time_range = f"{start_time}-{end_time}"

                    # Workers'ları birleştirerek bir string oluştur
                    worker_names = ", ".join([worker.get_name() for worker in workers])

                    # Atamayı listeye ekle
                    assignments.append({
                        "Product": product.get_serial_number(),
                        "Jig": jig.get_name(),
                        "Operation": operation.get_name(),
                        "Date": date_obj.date.strftime("%d.%m.%Y"),
                        "Shift": time_interval.shift,
                        "Time Interval": time_range,
                        "Workers": worker_names
                    })
        return assignments

    def debug(self):
        print("debug")

    def run_GUI(self):
        self.screenController = mainscreen.MainWindow()  # create GUI
        self.screenController.setMainController(self)
        print("GUI running")
        self.screenController.mainloop()


if __name__ == "__main__":
    main = MainController()
    main.run_GUI()
