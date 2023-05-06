class Project:
    sections = []

    def __init__(self, proj_id, name):
        self.id = proj_id
        self.name = name
        self.sections = []

    def id(self):
        return self.id

    def name(self):
        return self.name

    @property
    def all_sections(self):
        return self.sections

    def add_section(self, section):
        self.sections.append(section)
