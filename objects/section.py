class Section:
    def __init__(self, section_id, name):
        self.id = section_id
        self.name = name
        self.tasks = []

    def id(self):
        return self.id

    def name(self):
        return self.name

    @property
    def all_tasks(self):
        return self.tasks

    def add_task(self, task):
        self.tasks.append(task)
