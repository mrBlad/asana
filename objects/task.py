class Task:
    def __init__(self, task_id, name, date, status, compl_date, created_at):
        self.id = task_id
        self.name = name
        self.date = date # плановая(дедлайн)
        self.status = status
        self.compl_date = compl_date # дата выполнения(факт)
        self.created_at = created_at # дата создания вехи

    def id(self):
        return self.id

    def name(self):
        return self.name

    def date(self):
        """Плановая дата выполнения"""
        return self.date

    def status(self):
        return self.status

    def compl_date(self):
        """Дата выполнения(факт)"""
        return self.compl_date

    def created_at(self):
        """Дата создания"""
        return self.created_at

    @property
    def get_status(self):
        if self.status == 0:
            return u'Не выполнено'
        elif self.status == 1:
            return u'Выполнено'
        else:
            return ''
