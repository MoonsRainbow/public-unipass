import sys
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetExcelColTitleList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelBlList_GetCategory'
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            excel_col_title_list = []
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''SELECT admin_env_index, env_value FROM admin_env WHERE env_type = 1 AND flag_active = 1;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    excel_col_title_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                _query = '''SELECT admin_env_index, env_value FROM admin_env WHERE env_type = 2 AND flag_active = 1;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    excel_col_title_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                self.send(False, True, excel_col_title_list)
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


class AddExcelTitle(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _env_type: int, _env_value: str, _app_queue: [Queue, None] = None):
        self._code, self._env_type, self._env_value = 'ModelUserSetting_AddExcelTitle', _env_type, _env_value
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = f'''INSERT INTO admin_env(env_type, env_value) VALUES({self._env_type}, \'{self._env_value}\');'''
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


class DeleteExcelTitle(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _env_index: int, _app_queue: [Queue, None] = None):
        self._code, self._env_index = 'ModelUserSetting_DeleteExcelTitle', _env_index
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = f'''UPDATE admin_env SET flag_active = 0 WHERE admin_env_index = {self._env_index}'''
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

class SetDownloadPath(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _download_path: str, _app_queue: [Queue, None] = None):
        self._code, self._manager_index, self._download_path = 'ModelUserSetting_SetDownloadPath', _manager_index, _download_path
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT COUNT(*), manager_env_index 
                FROM manager_env 
                WHERE env_type = 1 AND manager_index = {}
                '''.format(self._manager_index)
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    if int(_der[1][0][0]) > 0:
                        _query = '''
                        UPDATE manager_env SET env_value = '{}' WHERE manager_env_index = {};
                        '''.format(self._download_path, _der[1][0][1])
                        _der2 = self._dbm.execute_query(_query)
                        if _der2[0]:
                            self.send(False, True)
                        else:
                            self.send(False, False, _der2[2])
                    else:
                        _query = '''
                        INSERT INTO manager_env(manager_index, env_type, env_value) VALUES({}, 1, '{}')
                        '''.format(self._manager_index, self._download_path)
                        _der2 = self._dbm.execute_query(_query)
                        if _der2[0]:
                            self.send(False, True)
                        else:
                            self.send(False, False, _der2[2])
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