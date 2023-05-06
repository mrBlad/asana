class Portfolio:
    def __init__(self, portf_id, name):
        self.id = portf_id
        self.name = name
        self.projects = []

    def id(self):
        return self.id

    def name(self):
        return self.name

    @property
    def all_projects(self):
        return self.projects

    def add_project(self, project):
        self.projects.append(project)
