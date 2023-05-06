import asyncio
from datetime import datetime

from dateutil import parser
import os

from asgiref.sync import sync_to_async
from xlwt import Workbook
import xlwt
from modules import arial10
from objects.frame import Frame
from objects.porfolio import Portfolio
from objects.project import Project
from objects.section import Section
from objects.task import Task
from objects.user import User
import PySimpleGUI as psg

import aiohttp as aiohttp


def main():
    async_loop = asyncio.get_event_loop()
    while True:
        authEvent, authValues = authWindow.frame().read()
        if authEvent == psg.WIN_CLOSED or authEvent == "Отмена":
            break

        authUser = User(None, None, None, None)

        if authEvent == "Войти":
            token = str(authValues["accessToken"])
            authUser = async_loop.run_until_complete(auth(token))
            if authUser.error():
                print("Exit")
                continue
            async_loop.run_until_complete(save_user(authUser))
        elif authEvent == "usersList" and authValues["usersList"] and authValues["usersList"][0] != " ":
            authUser = list(filter(lambda user: user.name() == str(authValues["usersList"][0]), upload_user()))[0]

        if authUser.name() is not None and authUser.token() is not None:

            authWindow.hide()

            portfolios, noArcheProj = async_loop.run_until_complete(upload_data(authUser))

            actual_portfolio = async_loop.run_until_complete(wrap(portfolios, noArcheProj))
            mainWindow = main_window(authUser.name(), [portfolio["portfolio_name"] for portfolio in actual_portfolio])

            mainWindow.show()
            while True:
                mainEvent, mainValues = mainWindow.frame().read()

                if mainEvent == psg.WIN_CLOSED or mainEvent == "Выйти":
                    break
                if mainEvent == "Выгрузить":
                    path = str(mainValues["pathToSave"])
                    name = str(mainValues["nameToSave"])
                    if mainValues["-PORTFOLIO-"]:
                        selected_portfs = list(
                            filter(lambda portfolio: portfolio["portfolio_name"] in mainValues["-PORTFOLIO-"],
                                   actual_portfolio))
                    else:
                        selected_portfs = actual_portfolio

                    portfolios_dict = async_loop.run_until_complete(create_hierarchy(selected_portfs))
                    async_loop.run_until_complete(export_excel(path, authUser, portfolios_dict, name))
                    print('done')

async def get_request(url, session, token, semaphore):
    async with semaphore:
        async with session.get(
                url=f'{url}',
                headers={'Authorization': "Bearer " + token}
        ) as response:
            pars = await response.json()
            # print(pars)
            return pars["data"]

# локальное сохранение данных пользователя
async def save_user(user):
    pathname = os.path.expanduser(os.path.join('~', 'Documents', 'AsanaUsers'))
    if not os.path.exists(pathname):
        os.mkdir(pathname)
    if not os.path.exists(os.path.join(pathname, user.gid() + '.txt')):
        textfile = open(os.path.join(pathname, user.gid() + '.txt'), 'w')
        lines = [user.gid() + '\n' + user.name() + '\n' + user.token() + '\n' + user.workspace_gid()]
        textfile.writelines(lines)
        textfile.close()

# для визуального отображения списка пользователей
def upload_user():
    pathname = os.path.expanduser(os.path.join('~', 'Documents', 'AsanaUsers'))
    if not os.path.exists(pathname):
        return False
    files = os.listdir(pathname)
    if not files:
        return False

    users = []

    for file in files:
        textfile = open(os.path.join(pathname, file), 'r')
        userData = textfile.readlines()
        users.append(
            User(
                gid=userData[0].rstrip(),
                name=userData[1].rstrip(),
                token=userData[2].rstrip(),
                workspace_id=userData[3].rstrip()
            )
        )
        textfile.close()
    return users

# запрос портфелий и их задач
async def get_request_portfolio_with_tasks(url, session, token, portfolio, semaphore):
    async with semaphore:
        async with session.get(
                url=f'{url}',
                headers={'Authorization': "Bearer " + token}
        ) as response:
            loop_data = await response.json()
            response_with_portfolio = {
                "gid": portfolio["gid"],
                "name": portfolio["name"],
                "tasks": loop_data["data"]
            }
            return response_with_portfolio

# авторизация пользователя
async def auth(token):
    async with aiohttp.ClientSession() as session:
        semaphore = asyncio.Semaphore(2)
        try:
            res = await get_request(
                url="https://app.asana.com/api/1.0/users/me?opt_fields=gid,name,workspaces.gid",
                session=session,
                token=token,
                semaphore=semaphore
            )

            # print(res)

            user = User(
                name=res["name"],
                gid=res["gid"],
                workspace_id=res["workspaces"][0]["gid"],
                token=token
            )

        except LookupError:
            print("Неверный токен-авторизации")
            user = User(
                name=None,
                gid=None,
                workspace_id=None,
                token=None,
                auth_error=True
            )

    print(user.token())

    return user

# окно авторизации пользователя
def auth_window(users):
    usernames = [user.name() for user in users] if users else []
    usernames.append(" ")

    left_layout = [
        [psg.Text("Выберите пользователя: ")],
        [psg.Listbox(values=usernames, size=(30, 5), auto_size_text=True, enable_events=True, key="usersList",
                     select_mode='LISTBOX_SELECT_MODE_SINGLE')]
    ]

    right_layout = [
        [psg.Text("Введите токен авторизации: ", auto_size_text=True)],
        [psg.InputText(key="accessToken")]
    ]

    layout = [
        [psg.Column(layout=left_layout, element_justification='l'),
         psg.Column(layout=right_layout, element_justification='с', expand_y=True)],
        [psg.Text(" ")],
        [psg.Column(layout=[[psg.Button("Отмена", size=(10, 1))]], element_justification='l'),
         psg.Column(layout=[[psg.Button("Войти", size=(10, 1))]], element_justification='right', expand_x=True)]
    ]
    return Frame("Анализ asana", layout, default_element_size=(40, 1))

# основное рабочее пространство
def main_window(username="", portfolios=[]):
    left_layout = [
        [psg.Text("Выберите портфели: ")],
        [psg.Listbox(values=portfolios, size=(30, 10), auto_size_text=True, enable_events=True, key="-PORTFOLIO-",
                     select_mode=psg.SELECT_MODE_EXTENDED)]
    ]

    right_layout = [
        [psg.Text("Укажите путь, где сохранить отчёт:", auto_size_text=True, justification='left')],
        [psg.InputText(size=(25, 1), justification='left', key="pathToSave"),
         psg.FolderBrowse("Выбрать", size=(10, 1))],
        [psg.Text("Укажите наименование отчёта:", auto_size_text=True, justification='left')],
        [psg.InputText(size=(25, 1), justification='left', key="nameToSave")],
    ]

    layout = [
        [psg.Text("Пользователь: " + username, auto_size_text=True, justification='left')],
        [psg.Column(layout=left_layout, element_justification='l'),
         psg.Column(layout=right_layout, element_justification='с', expand_y=True)],
        [psg.Text(" ")],
        [psg.Text(" ")],
        [psg.Column(layout=[[psg.Button("Выйти", size=(10, 1))]], element_justification='l'),
         psg.Column(layout=[[psg.Button("Выгрузить", size=(10, 1))]], key="loadDataInXML",
                    element_justification='right', expand_x=True)]
    ]
    return Frame("Анализ asana", layout, default_element_size=(40, 1))

# запрос данных о задачах
async def request_in_thread(milestonesPortfolios, index, portfolio, session, user, semaphore):
    milestonesPortfolios[index]["tasks"] = await asyncio.gather(*[asyncio.ensure_future(
        get_request(
            url="https://app.asana.com/api/1.0/tasks/"
                + str(task["gid"])
                + "?opt_fields=gid,name,completed,memberships.project.name,memberships.section.name,due_on,completed_at,created_at",
            session=session,
            token=user.token(),
            semaphore=semaphore
        )
    ) for task in portfolio["tasks"]], return_exceptions=True)
    # print(milestonesPortfolios[index]["tasks"])

# парсинг данных с asana
async def upload_data(user):
    async with aiohttp.ClientSession() as session:
        semaphore = asyncio.Semaphore(2)
        parsePortfoliosData = await get_request(
            url="https://app.asana.com/api/1.0/portfolios?workspace="
                + user.workspace_gid() + "&owner="
                + user.gid()
                + "&opt_fields=gid,name",
            session=session,
            token=user.token(),
            semaphore=semaphore
        )

        allNoArchivedProjects = await get_request(
            url="https://app.asana.com/api/1.0/workspaces/"
                + user.workspace_gid()
                + "/projects?archived=false&opt_fields=gid",
            token=user.token(),
            session=session,
            semaphore=semaphore
        )

        milestonesPortfolios = await asyncio.gather(*[asyncio.ensure_future(get_request_portfolio_with_tasks(
            url="https://app.asana.com/api/1.0/workspaces/"
                + user.workspace_gid()
                + "/tasks/search?resource_subtype=milestone"
                + "&portfolios.any=" + portfolio["gid"]
                + "&opt_fields=gid",
            token=user.token(),
            session=session,
            semaphore=semaphore,
            portfolio=portfolio
        )) for portfolio in parsePortfoliosData], return_exceptions=True)

        await asyncio.gather(
            *[asyncio.ensure_future(request_in_thread(milestonesPortfolios, index, portfolio, session, user, semaphore))
              for index, portfolio in enumerate(milestonesPortfolios)])

    return milestonesPortfolios, allNoArchivedProjects

# предобработка данных для отчистки некорректных данных
@sync_to_async
def wrap(tasks_arr, notArchivedResults):
    noEmptyTaskArr = [
        {
            "portfolio_gid": portfolio["gid"],
            "portfolio_name": portfolio["name"],
            "portfolio_tasks": list(filter(lambda task: task["memberships"], portfolio["tasks"]))
        }
        for portfolio in tasks_arr]

    noArchProj = list(map(lambda project: project['gid'], notArchivedResults))

    for portfolio_data in noEmptyTaskArr:
        for task_data in portfolio_data["portfolio_tasks"]:
            for member in task_data['memberships']:
                if member['project']:
                    if member['project']['gid'] not in noArchProj:
                        del member['project']
                        del member['section']

    return noEmptyTaskArr

# структуризация полученных данных
@sync_to_async
def create_hierarchy(tasks_arr):
    portfolios_dict = dict()
    all_portfolio_objs = {}  # объекты портфелей
    all_project_objs = {}  # проверочный словарь
    all_section_objs = {}  # проверочный словарь
    all_task_objs = {}  # проверочный словарь
    for portfolio_data in tasks_arr:
        for task_data in portfolio_data["portfolio_tasks"]:
            for member in task_data['memberships']:
                if member:
                    # print(member)
                    try:  # получаем проект или создаём
                        project = all_project_objs[member['project']['gid']]
                        all_project_objs.update({member['project']['gid']: project})
                    except KeyError:
                        project = Project(proj_id=member['project']['gid'], name=member['project']['name'])
                        all_project_objs.update({member['project']['gid']: project})
                    try:  # получаем секцию или создаём
                        section = all_section_objs[member['section']['gid']]
                        all_section_objs.update({member['section']['gid']: section})
                    except KeyError:
                        section = Section(section_id=member['section']['gid'], name=member['section']['name'])
                        all_section_objs.update({member['section']['gid']: section})
                    try:  # получаем веху или создаём
                        task = all_task_objs[task_data['gid']]
                        all_task_objs.update({task_data['gid']: task})
                    except KeyError:
                        task = Task(task_id=task_data['gid'], name=task_data['name'],
                                    status=task_data['completed'], date=task_data['due_on'],
                                    compl_date=task_data['completed_at'], created_at=task_data['created_at'])
                        all_task_objs.update({task_data['gid']: task})
                    try:  # если портфолио объект уже создан, просто добавим ему проект, иначе исключение и создание нового портфолио
                        portfolio = all_portfolio_objs[portfolio_data["portfolio_gid"]]
                        all_portfolio_objs.update({portfolio_data["portfolio_gid"]: portfolio})
                    except KeyError:
                        portfolio = Portfolio(portf_id=portfolio_data["portfolio_gid"],
                                              name=portfolio_data['portfolio_name'])
                        all_portfolio_objs.update({portfolio_data['portfolio_gid']: portfolio})

                    # связываем веху с секцией и секцию с проектом
                    if task not in section.all_tasks:
                        section.add_task(task)
                    if section not in project.all_sections:
                        project.add_section(section)
                    if project not in portfolio.all_projects:
                        portfolio.add_project(project)
                    portfolios_dict.update({portfolio.id: portfolio})
    # print(portfolios_dict)
    return portfolios_dict

# создание отчета по полученным данным в формате xlsx
@sync_to_async
def export_excel(path, authUser, portfolios_objs, filename="data"):
    wb = Workbook()
    sheet = wb.add_sheet('Данные')

    # style
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.bold = True
    style.font = font

    # first section row
    sheet.write(0, 0, 'Раздел', style=style)
    sheet.write(0, 1, 'Веха', style=style)
    width_date = 4050
    sheet.col(2).width = width_date
    style_date_word_wrap = xlwt.XFStyle()
    style_date_word_wrap.alignment.wrap = 1
    font = xlwt.Font()
    font.bold = True
    style_date_word_wrap.font = font
    sheet.write(0, 2, 'Дата планового выполнения', style=style_date_word_wrap)

    width_compl_date = 5000
    sheet.col(3).width = width_compl_date
    style_compl_date_word_wrap = xlwt.XFStyle()
    style_compl_date_word_wrap.alignment.wrap = 1
    font = xlwt.Font()
    font.bold = True
    style_compl_date_word_wrap.font = font
    sheet.write(0, 3, 'Дата фактического выполнения', style=style_compl_date_word_wrap)
    sheet.write(0, 4, 'Статус', style=style)

    # width_comment = 5350
    # sheet.col(5).width = width_comment
    style_comment_word_wrap = xlwt.XFStyle()
    style_comment_word_wrap.alignment.wrap = 1
    font = xlwt.Font()
    font.bold = True
    style_comment_word_wrap.font = font
    sheet.col(5).width = 3550 # 13
    sheet.write(0, 5, 'Комментарий', style=style_comment_word_wrap)

    # panes frozen
    sheet.set_panes_frozen(True)
    sheet.set_horz_split_pos(1)
    sheet.set_vert_split_pos(0)

    # row and col vars
    cell_row = 1
    cell_col = 0

    # data
    portfolios = portfolios_objs.values()

    for portfolio in portfolios:
        if portfolio:
            sheet.write(cell_row, cell_col, portfolio.name)

            cell_row += 2
            projects = portfolio.all_projects
            projects = sorted(projects, key=lambda project: project.id)
            # projects = list(reversed(projects))
            # print(type(projects))
            for project in projects:
                sheet.write_merge(cell_row, cell_row, 0, 4, project.name, xlwt.easyxf("align: horiz center"))
                sections = project.all_sections
                sections = sorted(sections, key=lambda section: section.id)
                # sections = list(reversed(sections))
                cell_row += 1
                for section in sections:
                    # width_sec = 5350  # 20
                    width_sec = 4050  # 15
                    sheet.col(cell_col).width = width_sec
                    # style for word wrap (перенос по словам)
                    style_sec_word_wrap = xlwt.XFStyle()
                    style_sec_word_wrap.alignment.wrap = 1
                    sheet.write(cell_row, 0, section.name, style=style_sec_word_wrap)
                    tasks = section.all_tasks
                    tasks = sorted(tasks, key=lambda task: parser.parse(task.created_at))
                    # tasks = sorted(tasks, key=lambda task: task.id)
                    # tasks = list(reversed(tasks))
                    for task in tasks:
                        # task name
                        # width_task = 15550  # 60
                        width_task = 7850  # 30
                        sheet.col(1).width = width_task
                        # style for word wrap (перенос по словам)
                        style_task_word_wrap = xlwt.XFStyle()
                        style_task_word_wrap.alignment.wrap = 1
                        sheet.write(cell_row, 1, task.name, style=style_task_word_wrap)
                        # task planned date
                        if task.date:
                            task_date_str = task.date
                            # width_task_date_str = arial10.fitwidth(task.date)
                            width_task_date_str = 3300 # 12
                            # if width_task_date_str > sheet.col(2).width:
                            sheet.col(2).width = width_task_date_str
                            task_date = datetime.strptime(task_date_str, '%Y-%m-%d')
                            task_date_str = datetime.strftime(task_date, '%d.%m.%Y')
                            sheet.write(cell_row, 2, task_date_str)
                        else:
                            # width_task_date = arial10.fitwidth("Нет даты")
                            width_task_date = 3300
                            # if width_task_date > sheet.col(2).width:
                            sheet.col(2).width = width_task_date
                            sheet.write(cell_row, 2, 'Нет даты')
                        # task fact date
                        if task.compl_date:
                            cutted_date_str = task.compl_date.split('T')[0]
                            # width_compl_date = arial10.fitwidth(cutted_date_str)
                            width_compl_date = 3550 # 13
                            # if width_compl_date > sheet.col(3).width:
                            sheet.col(3).width = width_compl_date
                            cutted_date = datetime.strptime(cutted_date_str, '%Y-%m-%d')
                            cutted_date_str = datetime.strftime(cutted_date, '%d.%m.%Y')
                            sheet.write(cell_row, 3, cutted_date_str)
                        else:
                            # width_compl_date = arial10.fitwidth("Нет даты")
                            width_compl_date = 3550
                            # if width_compl_date > sheet.col(3).width:
                            sheet.col(3).width = width_compl_date
                            sheet.write(cell_row, 3, "Нет даты")
                        # task status
                        # width_status = arial10.fitwidth(task.get_status)
                        width_status = 3210 # 12
                        # if width_status > sheet.col(4).width:
                        sheet.col(4).width = width_status
                        sheet.write(cell_row, 4, task.get_status)
                        # row end
                        cell_row += 1
                cell_row += 1
    # перенос в ячейке по словам в поле комментарий
    style_comment_field_word_wrap = xlwt.XFStyle()
    style_comment_field_word_wrap.alignment.wrap = 1
    for row in range(1, cell_row):
        sheet.write(row, 5, '', style=style_comment_field_word_wrap)
    # Коммент
    if not filename.strip():
        filename = "data"
    if not path:
        path = os.path.expanduser(os.path.join('~', 'Documents'))
    if os.path.isfile(path + '/' + filename.strip() + '.xls'):
        os.remove(path + '/' + filename.strip() + '.xls')
    wb.save(path + '/' + filename.strip() + '.xls')

if __name__ == '__main__':
    users = [] if not upload_user() else upload_user()
    authWindow = auth_window(users)
    authWindow.frame()
    main()
