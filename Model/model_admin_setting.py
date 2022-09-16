import sys
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetManagerList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelAdminSetting_GetManagerList'
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''SELECT manager_index, id, name FROM manager WHERE flag_active = 1 ORDER BY manager_index;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True, _der[1])
                else:
                    self.send(False, False, _der[2])
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()


class AddManager(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_info: dict, _app_queue: [Queue, None] = None):
        self._code, self._manager_info = 'ModelAdminSetting_AddManager', _manager_info
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''SELECT COUNT(*) FROM manager WHERE id = '{}';'''.format(self._manager_info['ID'])
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    if int(_der[1][0][0]) > 0:
                        self.send(False, True, False)
                    else:
                        _query = '''
                        INSERT INTO manager(id, pw, name) VALUES('{}', PASSWORD('{}'), '{}');
                        '''.format(self._manager_info['ID'], self._manager_info['PW'], self._manager_info['NAME'])
                        _der = self._dbm.execute_query(_query)
                        if _der[0]:
                            self.send(False, True, True)
                        else:
                            self.send(False, False, _der[2])
                else:
                    self.send(False, False, _der[2])
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()


class DeleteManager(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _app_queue: [Queue, None] = None):
        self._code, self._manager_index = 'ModelAdminSetting_DeleteManager', _manager_index
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                UPDATE manager SET flag_active = 0 WHERE manager_index = {}
                '''.format(self._manager_index)
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True)
                else:
                    self.send(False, False, _der[2])
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()


class ChangeManagerPassword(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _manager_password: str, _app_queue: [Queue, None] = None):
        self._code, self._manager_index, self._manager_password = 'ModelAdminSetting_ChangeManagerPassword', _manager_index, _manager_password
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                UPDATE manager SET pw = PASSWORD('{}') WHERE manager_index = {}
                '''.format(self._manager_password, self._manager_index)
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True)
                else:
                    self.send(False, False, _der[2])
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()
