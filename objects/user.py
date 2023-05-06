

class User:
    def __init__(self, name, token, gid, workspace_id, auth_error=False):
        self.__name = name
        self.__token = token
        self.__gid = gid
        self.__workspaceGID = workspace_id
        self.__authError = auth_error

    def gid(self):
        return self.__gid

    def workspace_gid(self):
        return self.__workspaceGID

    def name(self):
        return self.__name

    def token(self):
        return self.__token

    def error(self):
        return self.__authError
