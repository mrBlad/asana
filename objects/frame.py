import PySimpleGUI as psg

psg.theme('DarkAmber')


class Frame:
    def __init__(self, name, layout, **kwargs):
        self.__frame = psg.Window(name, layout, **kwargs)
        self.__active = False
        self.__name = ""

    def show(self):
        self.__active = True

    def active(self):
        return self.__active

    def hide(self):
        self.__active = False
        self.__frame.close()

    def name(self, name):
        self.__name = name

    def frame(self):
        return self.__frame