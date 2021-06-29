from functools import wraps
from tkRecruiting import NetworkError
import pyodbc


def monitor_network_state(method):
    """ Show error message in case of network error.
    """

    @wraps(method)
    def wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except pyodbc.Error as e:
            # Network error
            if e.args[0] in ('01000', '08S01', '08001'):
                NetworkError()

    return wrapper


class DBConnect(object):
    """ Provides connection to database and functions to work with server.
    """

    def __init__(self, *, server, db, uid, pwd):
        self._server = server
        self._db = db
        self._uid = uid
        self._pwd = pwd

    def __enter__(self):
        # Connection properties
        conn_str = (
            f'Driver={{SQL Server}};'
            f'Server={self._server};'
        )
        if self._db is not None:
            conn_str += f'Database={self._db};'
        if self._uid:
            conn_str += f'uid={self._uid};pwd={self._pwd}'
        else:
            conn_str += 'Trusted_Connection=yes;'
        self.__db = pyodbc.connect(conn_str)
        self.__cursor = self.__db.cursor()
        return self

    def __exit__(self, type, value, traceback):
        self.__db.close()

    @monitor_network_state
    def get_user_info(self, UserLogin):
        """ Check user permission.
            If access permitted returns True, otherwise None.
        """
        query = '''exec recruiting.user_info @UserLogin = ?'''

        self.__cursor.execute(query, UserLogin)
        isAccess = self.__cursor.fetchone()
        # check access to app
        if isAccess and (isAccess[2] == 1):
            return isAccess
        else:
            return None

    @monitor_network_state
    def update_log(self, UserLogin):
        """ Update logging table.
        """
        query = '''exec recruiting.usage_log @UserLogin = ?'''

        self.__cursor.execute(query, UserLogin)

    @monitor_network_state
    def create_request(self, userID, positionName, plannedDate, fileRequirements, commentText):
        """ Executes procedure that creates new request.
        """
        query = '''
        exec recruiting.create_vacancy @UserID = ?,
                                    @positionName = ?,
                                    @plannedDate = ?,
                                    @fileRequirements = ?,
                                    @commentText = ?
            '''
        # print(userID, positionName, plannedDate, fileRequirements, commentText)
        try:
            self.__cursor.execute(query, userID, positionName, plannedDate, fileRequirements, commentText)
            request_allowed = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_allowed
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def update_vacancy(self, id, modifiedUserID, responsibleID=None, fileCV=None,
                       statusID=None, startWork=None):
        """ Executes procedure that updates request.
        """
        query = '''
            exec recruiting.update_vacancy  @ID = ?,
                                            @modifiedID = ?,
                                            @responsibleID = ?,
                                            @fileCV = ?,
                                            @statusID = ?,
                                            @startWork = ?
                '''
        try:
            self.__cursor.execute(query, id, modifiedUserID, responsibleID,
                                  fileCV, statusID, startWork)
            request_allowed = self.__cursor.fetchone()[0]
            self.__db.commit()
            return request_allowed
        except pyodbc.ProgrammingError:
            return

    @monitor_network_state
    def get_offices(self):
        """ Returns list of available offices.
        """
        query = '''
        exec recruiting.get_offices
        '''
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_responsible(self, id=None):
        """ Returns list of active responsible users vacancy's list.
        """
        query = '''
        exec recruiting.get_responsible 0 , @ID = ?
        '''
        self.__cursor.execute(query, id)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_all_responsible(self, id=None):
        """ Returns list of available responsible HR users.
        """
        query = '''
        exec recruiting.get_responsible 1, @ID = ?
        '''
        self.__cursor.execute(query, id)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_current_responsible(self, id):
        """ Returns list of current responsible HR users.
        """
        query = '''
        exec recruiting.get_responsible 2, @ID = ?
        '''
        self.__cursor.execute(query, id)
        return self.__cursor.fetchall()

    @monitor_network_state
    def get_vacancies_list(self, statusID=None, officeID=None, responsibleID=None,
                           userOfficeID=None, userDepartmentID=None, isHR=None, UserID=None):
        """ Get list contracts with filters.
        """
        query = '''
           exec recruiting.get_vacancies_list @StatusID = ?,
                                              @OfficeID = ?,
                                              @ResponsibleUserID = ?,
                                              @userOfficeID = ?,
                                              @userDepartmentID = ?,
                                              @isHR = ?,
                                              @UserID = ?
           '''
        self.__cursor.execute(query, statusID, officeID, responsibleID, userOfficeID,
                              userDepartmentID, isHR, UserID)
        vacancies = self.__cursor.fetchall()
        self.__db.commit()
        return vacancies

    @monitor_network_state
    def get_status_list(self):
        """ Returns status list.
        """
        query = "exec recruiting.get_status_list"
        self.__cursor.execute(query)
        return self.__cursor.fetchall()

    @monitor_network_state
    def raw_query(self, query):
        """ Takes the query and returns output from db.
        """
        self.__cursor.execute(query)
        return self.__cursor.fetchall()


if __name__ == '__main__':
    with DBConnect(server='s-kv-center-s59', db='AnalyticReports',
                   uid='XXX', pwd='XXX') as sql:
        query = '''
                exec recruiting.get_responsible 1, @ID = 0
                '''
        print(sql.raw_query(query))
    print('Connected successfully.')
    input('Press Enter to exit...')
