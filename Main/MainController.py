from Screens import mainscreen
from Models.Dictionaries import SKILLS
from Models import Product
from Functions import SetCriticalOperation, WorkerAssigner
from Functions.ExcelDataLoader import ExcelDataLoader
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
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
            duration = operation.get_required_man_hours() / (7.5 * operation.get_required_worker())
            operation.set_operating_duration(duration)
            operation.set_remaining_duration(duration)  # Initialize remaining duration

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

        # Önce tüm tamamlanmış operasyonları bul
        completed_operations = []
        for operation in product.get_operations():
            if operation.get_completed():
                completed_operations.append(operation)

        # Her operasyon için tamamlanmış öncülleri güncelle
        for operation in product.get_operations():
            if not operation.get_completed():  # Sadece tamamlanmamış operasyonları güncelle
                # Öncül listelerini klonla
                uncompleted_predecessors = []

                # Tamamlanmamış öncülleri tespit et
                for pred in operation.get_predecessors():
                    if not pred.get_completed():
                        uncompleted_predecessors.append(pred)

                # Tamamlanmamış öncülleri güncelle
                operation.set_uncompleted_predecessors(uncompleted_predecessors)

                # Eğer artık tamamlanmamış öncül kalmadıysa, işaretleyelim
                if len(uncompleted_predecessors) == 0:
                    print(f"Operation {operation.get_name()} has no more incomplete predecessors")

    # CPM calculation
    def set_critical_operations(self, _sn):
        tasks = []
        product = self.get_product(_sn)
        calculator = SetCriticalOperation.Graph()

        # Sadece tamamlanmamış operasyonları CPM hesaplamasına dahil et
        for operation in product.get_operations():
            if not operation.get_completed():
                task = operation.get_name()
                # Operasyon süresini kontrol et, en az 0.01 olmasını sağla
                duration = max(operation.get_operating_duration(), 0.01)  # En az 0.01 süre

                # Sadece tamamlanmamış öncülleri ekle
                dependencies = []
                for op in operation.get_uncompleted_predecessors():
                    dependencies.append(op.get_name())

                # Grafiğe ekle
                calculator.add_task(task, duration, dependencies)
                tasks.append(task)

        # Kritik operasyonları hesapla
        critical_operations, earliest_start, latest_finish = calculator.find_critical_operations()

        print("Kritik Operasyonlar:", critical_operations)
        print("En Erken Başlama Zamanları:", earliest_start)
        print("En Geç Tamamlanma Zamanları:", latest_finish)

        # Kritik operasyonları listeye ekle
        critical_op_obj_list = []
        for op_name in critical_operations:
            op_obj = product.get_operation_by_name(op_name)
            if op_obj and not op_obj.get_completed():  # Tamamlanmamış olduğundan emin ol
                op_obj.set_early_start(earliest_start[op_name])
                op_obj.set_late_finish(latest_finish[op_name])
                if op_obj.get_early_start() == 0:
                    critical_op_obj_list.append(op_obj)

        # Kritik operasyonları ürüne ata
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
                # Öncülleri tamamlanmış ve kendisi henüz tamamlanmamış operasyonları seç
                if not op.get_completed() and len(op.get_uncompleted_predecessors()) == 0:
                    critical_ops.append((product, op))
                    print(f"Critical operation: {op.get_name()} for product {product.get_serial_number()}")

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

    def previous_operation_control(self, operation, time_interval):
        """
        Checks if an operation can be scheduled in the given time interval
        by verifying that all predecessors are completed and not running in the same interval.

        Args:
            operation: The operation to be scheduled
            time_interval: The time interval to check

        Returns:
            True if the operation can be scheduled, False otherwise
        """
        # First check if all predecessors are completed
        for pred in operation.get_predecessors():
            if not pred.get_completed():
                return False  # Cannot schedule if any predecessor is not completed

        # Now check if any predecessor is scheduled to run in this interval
        assignments = time_interval.get_assignments()

        for assignment in assignments:
            if len(assignment) >= 3:  # Make sure assignment has enough elements
                assigned_jig, assigned_product, assigned_operation, assigned_workers = assignment

                # Check if the assigned operation is a predecessor of our operation
                for prev_op in operation.get_previous_operations():
                    if assigned_operation.get_name() == prev_op.get_name():
                        return False  # Predecessor is running in this interval

                # Also check that we're not trying to schedule two operations that require the same predecessors
                # (to prevent logical conflicts in scheduling)
                for pred in assigned_operation.get_predecessors():
                    for our_pred in operation.get_predecessors():
                        if pred.get_name() == our_pred.get_name():
                            # If we share predecessors and neither operation is completed,
                            # they might be competing for the same resources
                            if not assigned_operation.get_completed() and not operation.get_completed():
                                return False

        return True  # All checks passed, operation can be scheduled

    def initiate_assignment(self, critical_op_list):
        """
        Initiates the assignment process for critical operations.
        Handles deadlock detection and supports partial assignment of operations.

        Args:
            critical_op_list: List of tuples (product, operation) to be scheduled
        """
        # Compare to previous operation list - if the same, we're stuck
        if critical_op_list == self.__critical_op_check_list and critical_op_list:
            print("Same operation list detected - forcing resolution")

            # Try a force assignment for the first operation in the list
            product, operation = critical_op_list[0]
            print(f"Forcing assignment for operation {operation.get_name()} of product {product.get_serial_number()}")

            # Get earliest available time interval
            intervals_list = self.get_ScheduleObject().get_sorted_time_intervals()
            if intervals_list:
                # Force assign to earliest available interval with any available workers
                for interval in intervals_list:
                    # Select best workers for assignment
                    selected_workers = self.select_best_workers_for_assignment(operation, interval)

                    if selected_workers and len(selected_workers) >= operation.get_required_worker():
                        # Force the assignment regardless of other constraints
                        jig = product.get_current_jig()

                        # If jig is incompatible, try to change it
                        if not self.jig_compatibility_control(product, operation):
                            self.change_jig(product, operation)
                            jig = product.get_current_jig()

                        # Create the assignment and make sure operation is marked as completed
                        self.create_assignment(interval, jig, product, operation, selected_workers)
                        operation.set_completed(True)  # Double-check it's marked as completed

                        # Update dependencies
                        product_sn = product.get_serial_number()
                        self.remove_completed_predecessors(product_sn)

                        print(f"Forced assignment successful for operation {operation.get_name()}")

                        # Reset critical ops check list to break comparison cycle
                        self.__critical_op_check_list = []
                        # Continue with the next operation
                        self.make_assignment_preparetions()
                        return

                # If no assignment could be made for first operation, remove it and continue
                print(f"Could not force assignment for operation {operation.get_name()}, removing from critical path")
                # Manually remove this operation from consideration
                operation.set_completed(True)  # Mark as completed to remove from critical path

                # Reset critical ops check list to break comparison cycle
                self.__critical_op_check_list = []
                self.make_assignment_preparetions()
                return

        # Store current operation list for comparison in next iteration
        self.__critical_op_check_list = critical_op_list

        # Process each critical operation
        for product, operation in critical_op_list:
            intervals_list = self.get_ScheduleObject().get_sorted_time_intervals()

            # 1. Find the latest finish time of predecessors
            latest_finish_time = self.find_latest_finish_time_of_predecessors(operation)

            # If latest_finish_time is None, it means some predecessors haven't been scheduled yet
            if latest_finish_time is None:
                print(
                    f"Skipping operation {operation.get_name()} for product {product.get_serial_number()} - predecessors not yet scheduled")
                continue  # Skip this operation for now

            # 2. Filter intervals that start after the latest finish time
            filtered_intervals = self.filter_intervals_after_time(intervals_list, latest_finish_time)

            if not filtered_intervals:
                print(
                    f"No valid intervals for operation {operation.get_name()} after predecessor completion time {latest_finish_time}")
                continue  # No valid intervals available, skip this operation

            # Initialize variables for tracking assignments
            remaining_duration = operation.get_operating_duration()
            successfully_assigned = False

            # Check scheduling options: first try with same product, then without same product
            for same_product_mode in [True, False]:
                if successfully_assigned:
                    break

                for interval in filtered_intervals:
                    # Skip if a previous operation is still running in this interval
                    if not self.previous_operation_control(operation, interval):
                        continue

                    # Skip this interval if we're in same_product_mode but no same product in interval,
                    # or if we're not in same_product_mode but there is same product in interval
                    has_same_product = self.same_product_control(product, interval)
                    if same_product_mode != has_same_product:
                        continue

                    # If same product is already in interval, check jig capacity
                    if has_same_product and not self.check_jig_capacity(product, operation, interval):
                        continue

                    # Verify we have enough workers
                    if not self.compatible_worker_number_check(operation, interval):
                        continue

                    # Handle jig compatibility
                    current_jig = product.get_current_jig()
                    if not has_same_product:  # Only change jig if not already using one for same product
                        if not self.jig_compatibility_control(product, operation):
                            self.change_jig(product, operation)

                    # Get workers for this assignment
                    selected_workers = self.select_best_workers_for_assignment(operation, interval)
                    if not selected_workers:
                        continue

                    # Verify we can schedule the complete duration
                    jig = product.get_current_jig()
                    assignment_intervals = []
                    current_interval = interval
                    assignment_duration = 0

                    # Try to find consecutive intervals for the full operation
                    while assignment_duration < remaining_duration:
                        # Verify this interval passes all constraints
                        if (not self.previous_operation_control(operation, current_interval) or
                                (self.same_product_control(product, current_interval) and
                                 not self.check_jig_capacity(product, operation, current_interval)) or
                                not self.compatible_worker_number_check(operation, current_interval)):
                            break

                        assignment_intervals.append(current_interval)
                        assignment_duration += 0.25  # Each interval is 0.25 hours

                        # Look for next interval
                        next_interval = self.get_next_interval(current_interval, filtered_intervals)
                        if next_interval is None:
                            break
                        current_interval = next_interval

                    # Check if we found enough intervals for full assignment
                    if assignment_duration >= remaining_duration:
                        # Create assignments for each interval
                        for assigned_interval in assignment_intervals:
                            self.create_assignment(assigned_interval, jig, product, operation, selected_workers)

                        # Mark operation as completed since full duration is scheduled
                        operation.set_completed(True)
                        successfully_assigned = True
                        break  # Break the interval loop

                    # If we can't schedule the complete operation, but we found some intervals,
                    # still create partial assignments - this is a key improvement
                    elif assignment_intervals and assignment_duration > 0:
                        print(
                            f"Partial assignment for operation {operation.get_name()}, scheduled {assignment_duration} of {remaining_duration} hours")

                        # Create assignments for available intervals
                        for assigned_interval in assignment_intervals:
                            self.create_assignment(assigned_interval, jig, product, operation, selected_workers)

                        # Track the remaining duration needed
                        remaining_duration -= assignment_duration

                        # Don't mark as completed yet, but consider this a successful partial assignment
                        successfully_assigned = True
                        break  # Break the interval loop to try other operations

            # If we couldn't assign this operation at all, log it
            if not successfully_assigned:
                print(f"Could not assign operation {operation.get_name()} for product {product.get_serial_number()}")

        # Continue with preparation for the next round of assignments
        self.make_assignment_preparetions()

    def get_next_interval(self, current_interval, intervals_list):
        """
        Mevcut interval'dan sonraki interval'ı bulur.
        """
        current_index = intervals_list.index(current_interval)
        if current_index + 1 < len(intervals_list):
            return intervals_list[current_index + 1]
        return None  # Sonraki interval yok

    def select_best_workers_for_assignment(self, operation, time_interval):
        """
        Operasyon için en uygun işçileri seçer:
        1. Vardiyası uyumlu olan işçileri filtreler
        2. Gerekli becerilere sahip işçileri filtreler
        3. En az atama yapılmış işçileri önceliklendirip seçer
        """
        required_skills = operation.get_required_skills()
        required_worker_count = operation.get_required_worker()

        # Uygun beceri ve vardiyaya sahip tüm işçileri bul
        qualified_workers = []

        for worker in time_interval.available_workers:
            # Beceri kontrolü
            worker_skills = worker.get_skills()
            is_qualified = False

            if worker_skills in SKILLS:
                worker_skill_set = SKILLS[worker_skills]
                if required_skills in worker_skill_set:
                    is_qualified = True
            else:
                if required_skills == worker_skills:
                    is_qualified = True

            if is_qualified:
                # Vardiya kontrolü
                worker_shift_schedule = worker.get_shift_schedule()
                for schedule_entry in worker_shift_schedule:
                    schedule_date, schedule_shift, available_hours = schedule_entry

                    # Tarih ve vardiya kontrolü
                    if schedule_date == time_interval.get_date() and schedule_shift == time_interval.get_shift():
                        # İşçi hem beceri hem vardiya açısından uygun
                        qualified_workers.append(worker)
                        break

        # Yeterli sayıda uygun işçi var mı kontrol et
        if len(qualified_workers) < required_worker_count:
            return None

        # İşçilerin atama sayılarını kontrol et, eğer atama sayısı tanımlı değilse 0 olarak başlat
        for worker in qualified_workers:
            if not hasattr(worker, 'assignment_count'):
                worker.assignment_count = 0

        # İşçileri atama sayısına göre sırala (en az atama yapılan önce)
        sorted_workers = sorted(qualified_workers, key=lambda w: w.assignment_count)

        # Gerekli sayıda işçiyi seç
        selected_workers = sorted_workers[:required_worker_count]

        return selected_workers

    def find_latest_finish_time_of_predecessors(self, operation):
        op = operation
        latest_finish_time = None
        intervals = self.get_ScheduleObject().get_sorted_time_intervals()

        # If no predecessors, can start at the earliest possible time
        if len(op.get_previous_operations()) == 0:
            first_interval = intervals[0]
            latest_finish_time = (first_interval.get_date(), first_interval.interval[0])
            return latest_finish_time

        # Check if all predecessors have end times
        all_predecessors_have_end_time = True
        for prev_op in op.get_previous_operations():
            if not prev_op.get_end_datetime() and not prev_op.get_completed():
                all_predecessors_have_end_time = False
                break

        # If any predecessor doesn't have an end time, this operation can't be scheduled yet
        if not all_predecessors_have_end_time:
            return None

        # Find the latest end time among all predecessors
        for prev_op in op.get_previous_operations():
            # Skip if the predecessor is already marked as completed but doesn't have end_datetime
            if prev_op.get_completed() and not prev_op.get_end_datetime():
                continue

            # If predecessor has an end_datetime, compare it with current latest
            if prev_op.get_end_datetime():
                if latest_finish_time is None:
                    latest_finish_time = prev_op.get_end_datetime()
                elif prev_op.get_end_datetime()[0] > latest_finish_time[0]:
                    latest_finish_time = prev_op.get_end_datetime()
                elif (prev_op.get_end_datetime()[0] == latest_finish_time[0] and
                      prev_op.get_end_datetime()[1] > latest_finish_time[1]):
                    latest_finish_time = prev_op.get_end_datetime()

        return latest_finish_time

    # 2. Fix the filter_intervals_after_time method to ensure operations start strictly after predecessors end
    def filter_intervals_after_time(self, intervals_list, start_time):
        if start_time is None:
            return []  # If no start time is found, don't schedule (empty list)

        filtered_intervals = []
        for interval in intervals_list:
            interval_date = interval.get_date()
            interval_start_time = interval.interval[0]  # Interval's start time

            # Only include intervals that start strictly after the predecessor end time
            if interval_date > start_time[0]:
                filtered_intervals.append(interval)
            elif interval_date == start_time[0] and interval_start_time > start_time[1]:
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

        # İşçilerin atama sayısını artır
        for worker in workers:
            if not hasattr(worker, 'assignment_count'):
                worker.assignment_count = 0
            worker.assignment_count += 1

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

        # Ürünlerin ilerlemesini hesapla ve öncülleri güncelle
        for product in plist:
            sn = product.get_serial_number()
            self.calculate_product_progress(sn)
            self.remove_completed_predecessors(sn)
            self.set_critical_operations(sn)

        # Operasyonları süreye göre sırala
        self.sort_operations_by_duration()
        # Ürünleri ilerleme durumuna göre sırala
        self.sort_products_by_progress()

        # Yeni kritik operasyon listesi
        new_critical_ops = self.get_all_critical_operations()

        # Eğer kritik operasyon listesi boşsa, işlem tamamlandı demektir
        if not new_critical_ops:
            print("Assignment complete - no more critical operations.")
            return

        # Sonsuz döngü kontrolü için kritik operasyonları yazdır
        print("Critical operations for next assignment:")
        for product, op in new_critical_ops:
            print(f"Product: {product.get_serial_number()}, Operation: {op.get_name()}")

        # Atama işlemini başlat
        self.initiate_assignment(new_critical_ops)
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

    def export_assignments_to_excel(self, file_path=None):
        """
        Atama çıktılarını Excel dosyasına aktarır.
        Her ürün için ayrı bir sayfa oluşturur.

        :param file_path: Excel dosyasının kaydedileceği konum. None ise kullanıcıdan sorulur.
        :return: Başarılı olursa True, aksi halde False
        """
        try:
            # Atama çıktılarını al
            assignments = self.get_assignments_for_output()

            if not assignments:
                print("No assignments to export.")
                return False

            # Eğer dosya yolu belirtilmemişse, kullanıcıdan al
            if not file_path:
                from tkinter import filedialog
                file_path = filedialog.asksaveasfilename(
                    title="Save Excel File",
                    filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
                    defaultextension=".xlsx"
                )

                if not file_path:  # Kullanıcı iptal ettiyse
                    return False

            # Yeni bir Excel çalışma kitabı oluştur
            wb = openpyxl.Workbook()
            # Varsayılan sayfayı sil
            wb.remove(wb.active)

            # Ürünlere göre atamaları grupla
            product_assignments = {}
            for assignment in assignments:
                product_serial = assignment["Product"]
                if product_serial not in product_assignments:
                    product_assignments[product_serial] = []
                product_assignments[product_serial].append(assignment)

            # Her ürün için ayrı bir sayfa oluştur
            for product_serial, product_assignments_list in product_assignments.items():
                # Operasyon adına göre sırala
                sorted_assignments = sorted(product_assignments_list,
                                            key=lambda x: int(x["Operation"]) if x["Operation"].isdigit() else x[
                                                "Operation"])

                # Yeni sayfa oluştur
                sheet = wb.create_sheet(f"Product {product_serial}")

                # Başlıkları ayarla
                headers = ["Operation", "Jig", "Date", "Shift", "Time Interval", "Workers"]
                for col_idx, header in enumerate(headers, 1):
                    cell = sheet.cell(row=1, column=col_idx)
                    cell.value = header
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                    cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

                    # Sütun genişliklerini ayarla
                    if header == "Workers":
                        sheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 30
                    elif header == "Time Interval":
                        sheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15
                    else:
                        sheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 12

                # Verileri doldur
                for row_idx, assignment in enumerate(sorted_assignments, 2):
                    sheet.cell(row=row_idx, column=1).value = assignment["Operation"]
                    sheet.cell(row=row_idx, column=2).value = assignment["Jig"]
                    sheet.cell(row=row_idx, column=3).value = assignment["Date"]
                    sheet.cell(row=row_idx, column=4).value = assignment["Shift"]
                    sheet.cell(row=row_idx, column=5).value = assignment["Time Interval"]
                    sheet.cell(row=row_idx, column=6).value = assignment["Workers"]

                    # Hücre stillerini ayarla
                    for col_idx in range(1, 7):
                        cell = sheet.cell(row=row_idx, column=col_idx)
                        cell.alignment = Alignment(horizontal='center')
                        if row_idx % 2 == 0:
                            cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

                # Özet bilgiler
                row_idx = len(sorted_assignments) + 3
                sheet.cell(row=row_idx, column=1).value = "Total Operations:"
                sheet.cell(row=row_idx, column=2).value = len(sorted_assignments)
                sheet.cell(row=row_idx, column=1).font = Font(bold=True)

                # Sayfayı güzelleştir
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )

                for row in sheet.iter_rows(min_row=1, max_row=row_idx, min_col=1, max_col=6):
                    for cell in row:
                        cell.border = thin_border

            # Özet sayfası ekle
            summary_sheet = wb.create_sheet("Summary", 0)  # İlk sayfaya yerleştir
            summary_sheet.cell(row=1, column=1).value = "Assignment Summary"
            summary_sheet.cell(row=1, column=1).font = Font(bold=True, size=14)
            summary_sheet.cell(row=1, column=1).alignment = Alignment(horizontal='center')
            summary_sheet.merge_cells('A1:D1')

            summary_sheet.cell(row=3, column=1).value = "Product"
            summary_sheet.cell(row=3, column=2).value = "Operations Count"
            summary_sheet.cell(row=3, column=3).value = "First Date"
            summary_sheet.cell(row=3, column=4).value = "Last Date"

            for col_idx in range(1, 5):
                summary_sheet.cell(row=3, column=col_idx).font = Font(bold=True)
                summary_sheet.cell(row=3, column=col_idx).alignment = Alignment(horizontal='center')
                summary_sheet.cell(row=3, column=col_idx).fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7",
                                                                             fill_type="solid")
                summary_sheet.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 15

            row_idx = 4
            for product_serial, product_assignments_list in product_assignments.items():
                first_date = min(product_assignments_list, key=lambda x: datetime.strptime(x["Date"], "%d.%m.%Y"))[
                    "Date"]
                last_date = max(product_assignments_list, key=lambda x: datetime.strptime(x["Date"], "%d.%m.%Y"))[
                    "Date"]

                summary_sheet.cell(row=row_idx, column=1).value = product_serial
                summary_sheet.cell(row=row_idx, column=2).value = len(product_assignments_list)
                summary_sheet.cell(row=row_idx, column=3).value = first_date
                summary_sheet.cell(row=row_idx, column=4).value = last_date

                # Hücre stillerini ayarla
                for col_idx in range(1, 5):
                    cell = summary_sheet.cell(row=row_idx, column=col_idx)
                    cell.alignment = Alignment(horizontal='center')
                    if row_idx % 2 == 0:
                        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    cell.border = thin_border

                row_idx += 1

            # Dosyayı kaydet
            wb.save(file_path)
            print(f"Assignments exported to {file_path}")
            return True

        except Exception as e:
            print(f"Error exporting assignments to Excel: {e}")
            import traceback
            traceback.print_exc()
            return False
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
