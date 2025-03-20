class Task:
    def __init__(self, name, duration, dependencies):
        self.name = name
        self.duration = duration
        self.dependencies = dependencies
        self.early_start = 0
        self.early_finish = 0
        self.late_start = float('inf')
        self.late_finish = float('inf')
        self.total_float = 0


def calculate_cpm(tasks):
    task_dict = {task.name: task for task in tasks}

    # =============================================================================
    # FORWARD PASS
    # =============================================================================
    for task in tasks:
        task.early_start = max(
            [task_dict[dep].early_finish for dep in task.dependencies] or [0]
        )
        task.early_finish = task.early_start + task.duration

    # =============================================================================
    # BACKWARD PASS
    # =============================================================================
    max_finish = max(task.early_finish for task in tasks)
    for task in reversed(tasks):
        task.late_finish = min(
            [task_dict[dep].late_start for dep in task_dict if task.name in task_dict[dep].dependencies] or [max_finish]
        )
        task.late_start = task.late_finish - task.duration
        task.total_float = task.late_start - task.early_start

    critical_operations = [task.name for task in tasks if task.total_float == 0]

    return tasks, critical_operations


if __name__ == "__main__":
    tasks = [
        Task("A", 7),
        Task("B", 9),
        Task("C", 12, ["A"]),
        Task("D", 8, ["A", "B"]),
        Task("E", 9, ["D"]),
        Task("F", 6, ["C", "E"]),
        Task("G", 5, ["E"]),
    ]

    tasks, critical_path = calculate_cpm(tasks)

    print("Tasks and durations:")
    for task in tasks:
        print(
            f"Task: {task.name}, Expt: {task.duration}, ES: {task.early_start}, EF: {task.early_finish}, "
            f"LS: {task.late_start}, LF: {task.late_finish}, Float: {task.total_float}"
        )

    print(f"Critical Path: {' -> '.join(critical_path)}")
