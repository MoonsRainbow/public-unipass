import sys
import base64
import configparser
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetLocalData(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _local_file_path: str, _key: str, _app_queue: [Queue, None] = None):
        self._code, self._local_file_path, self._config, self._key = 'ModelLogin_GetLocalSavedInfo', _local_file_path, configparser.ConfigParser(), _key
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            self._config.read(self._local_file_path, 'utf-8')
            _dec = []
            _enc_id = base64.urlsafe_b64decode(str(self._config['ACCOUNT']['ID']))
            for i in range(len(_enc_id)):
                _key_char = self._key[i % len(self._key)]
                _dec_char = chr((256 + _enc_id[i] - ord(_key_char)) % 256)
                _dec.append(_dec_char)
            _dec_id = ''.join(_dec) if ''.join(_dec) != 'Deleted' else ''
            _dec = []
            _enc_pw = base64.urlsafe_b64decode(str(self._config['ACCOUNT']['PW']))
            for i in range(len(_enc_pw)):
                _key_char = self._key[i % len(self._key)]
                _dec_char = chr((256 + _enc_pw[i] - ord(_key_char)) % 256)
                _dec.append(_dec_char)
            _dec_pw = ''.join(_dec) if ''.join(_dec) != 'Deleted' else ''
            self.send(False, True, {
                'ID': _dec_id,
                'PW': _dec_pw
            })
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._config.clear()


class SetLocalData(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _local_file_path: str, _key: str, _login_info: dict, _app_queue: [Queue, None] = None):
        self._code, self._local_file_path, self._config, self._key, self._login_info = 'ModelLogin_SetLocalSavedInfo', _local_file_path, configparser.ConfigParser(), _key, _login_info
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            for k, v in self._login_info.items():
                _enc = []
                for i in range(len(v)):
                    _key_char = self._key[i % len(self._key)]
                    _enc_char = (ord(v[i]) + ord(_key_char)) % 256
                    _enc.append(_enc_char)
                _enc = str(base64.urlsafe_b64encode(bytes(_enc)))[2:-1]
                self._config.read(self._local_file_path)
                for sn in self._config.sections():
                    for vn in self._config.options(sn):
                        self._config.set(sn, vn, self._config[sn][vn])
                if not self._config.has_section('ACCOUNT'):
                    self._config.add_section('ACCOUNT')
                self._config.set('ACCOUNT', k, _enc)
                with open(self._local_file_path, "w") as conf_writer:
                    self._config.write(conf_writer)
                self._config.clear()
            self.send(False, True)
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._config.clear()


class GetAuthority(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _login_info: dict, _app_queue: [Queue, None] = None):
        self._code, self._login_info = 'ModelLogin_GetAuthority', _login_info
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT COUNT(*), m.manager_index, name, IFNULL(me.env_value, '') AS env_path, flag_admin
                FROM manager m LEFT JOIN manager_env me on m.manager_index = me.manager_index AND me.env_type = 1
                WHERE flag_active = 1
                  AND id = '{}'
                  AND pw = PASSWORD('{}');
                '''.format(self._login_info['ID'], self._login_info['PW'])
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
