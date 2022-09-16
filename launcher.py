import os
import sys
import psutil
import pymysql
import zipfile
import tkinter
import datetime
import requests
import threading
from queue import Queue
from tkinter import messagebox
from win32com.shell import shell


class Launcher(tkinter.Tk):
    def __init__(self):
        try:
            super().__init__()
            self.Q = Queue()
            self.CT = None
            self.ENV_ROOT = None
            self.ENV_DIR = '/UNIPASS-AGL'
            self.CON_DIR = '/Config'
            self.ICON_FILE = '/icon.ico'
            self.CON_FILE = '/qdd08j6.ini'
            self.overrideredirect(True)
            self.VER_NO = None
            self.VER_LINK = None
            self.geometry('{}x{}+{}+{}'.format(
                300,
                100,
                int(self.winfo_screenwidth() / 2) - int(300 / 2),
                int(self.winfo_screenheight() / 2) - int(100 / 2)
            ))

            def update_status(_top, _status):
                if _status >= 10:
                    _status = 0
                else:
                    _status += 1

                try:
                    label_status.configure(text=' ● ' * _status + ' ○ ' * (10 - _status))
                    _top.after(100, update_status, _top, _status)
                except tkinter.TclError:
                    pass

            label_status = tkinter.Label(text=' ○  ○  ○  ○  ○  ○  ○  ○  ○  ○ ', font=('Consolas', 10))
            label_update = tkinter.Label(self, textvariable=tkinter.StringVar(master=self, name='MSG', value='런쳐를 실행 하는 중 입니다 ...'), font=('Consolas', 10))
            label_status.pack(anchor=tkinter.S, expand=tkinter.TRUE)
            label_update.pack(anchor=tkinter.N, expand=tkinter.TRUE)
            self.after(100, update_status, self, 1)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def check_env_root(self):
        try:
            for k, v in enumerate(os.environ):
                if v == 'PROGRAMFILES' or v == 'PROGRAMFILES(X86)':
                    self.ENV_ROOT = os.environ.get(v).replace('\\', '/')
                    break
            if not os.path.isdir(self.ENV_ROOT + self.ENV_DIR):
                os.mkdir(self.ENV_ROOT + self.ENV_DIR)
            if not os.path.isdir(self.ENV_ROOT + self.ENV_DIR + self.CON_DIR):
                os.mkdir(self.ENV_ROOT + self.ENV_DIR + self.CON_DIR)
            if not os.path.isfile(self.ENV_ROOT + self.ENV_DIR + self.CON_DIR + self.CON_FILE):
                with open(self.ENV_ROOT + self.ENV_DIR + self.CON_DIR + self.CON_FILE, 'w') as ini:
                    ini.write('[ACCOUNT]\n')
                    ini.write('id = \n')
                    ini.write('pw = \n')

            self.check_run_status()
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def check_run_status(self, is_run=False):
        try:
            if is_run:
                if self.Q.empty():
                    self.after(100, self.check_run_status, True)
                else:
                    response = self.Q.get()
                    if response[0]:
                        if response[1]:
                            is_restart = messagebox.askyesno(master=self, title='UNIPASS ::', message='이미 UNIPASS 가 실행 중 입니다.\n종료하고 다시 실행할까요?')
                            if is_restart == tkinter.YES:
                                _ver_list = [_exe for _exe in os.listdir(self.ENV_ROOT + self.ENV_DIR) if '.exe' in str(_exe)]
                                for proc in psutil.process_iter():
                                    if proc.name() in _ver_list:
                                        proc.kill()
                                self.request_version_info()
                            else:
                                self.destroy()
                        else:
                            self.setvar('MSG', '업데이트를 확인 하는 중 입니다 ...')
                            self.request_version_info()
                    else:
                        self.record_log(response[1])
                        self.destroy()
            else:
                self.CT = CheckRunStatus(
                    _q=self.Q,
                    _ep=self.ENV_ROOT + self.ENV_DIR
                )
                self.CT.setDaemon(True)
                self.CT.start()
                self.after(100, self.check_run_status, True)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def check_version(self):
        try:
            if not os.path.isfile(self.ENV_ROOT + self.ENV_DIR + '/{}.exe'.format(self.VER_NO)):
                self.setvar('MSG', '새로운 버전을 다운로드 중 입니다 ...')
                self.request_download()
            else:
                self.run()
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def run(self):
        try:
            os.popen(self.ENV_ROOT + self.ENV_DIR + '/{}.exe'.format(self.VER_NO))
            self.setvar('MSG', 'UNIPASS 를 실행 중 입니다 ...')
            self.after(5000, self.destroy)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def request_zip_extract(self, is_run=False):
        try:
            if is_run:
                if self.Q.empty():
                    self.after(100, self.request_zip_extract, True)
                else:
                    response = self.Q.get()
                    if response[0]:
                        self.run()
                    else:
                        if 'OperationalError' in str(response[1][1]):
                            messagebox.showerror(master=self, title='UNIPASS ::', message='인터넷 연결을 확인해주세요.')
                        self.record_log(response[1])
                        self.destroy()
            else:
                self.CT = ZipExtract(
                    _q=self.Q,
                    _ep=self.ENV_ROOT + self.ENV_DIR + '/{}.zip'.format(self.VER_NO),
                    _vn=self.VER_NO
                )
                self.CT.setDaemon(True)
                self.CT.start()
                self.after(100, self.request_zip_extract, True)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def request_download(self, is_run=False):
        try:
            if is_run:
                if self.Q.empty():
                    self.after(100, self.request_download, True)
                else:
                    response = self.Q.get()
                    if response[0]:
                        self.setvar('MSG', '압축을 해제 중 입니다 ...')
                        self.after(3000, self.request_zip_extract)
                    else:
                        if 'OperationalError' in str(response[1][1]):
                            messagebox.showerror(master=self, title='UNIPASS ::', message='인터넷 연결을 확인해주세요.')
                        self.record_log(response[1])
                        self.destroy()
            else:
                self.CT = Download(
                    _q=self.Q,
                    _ep=self.ENV_ROOT + self.ENV_DIR + '/{}.zip'.format(self.VER_NO),
                    _vn=self.VER_NO,
                    _vl=self.VER_LINK)
                self.CT.setDaemon(True)
                self.CT.start()
                self.after(100, self.request_download, True)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    def request_version_info(self, is_run=False):
        try:
            if is_run:
                if self.Q.empty():
                    self.after(100, self.request_version_info, True)
                else:
                    response = self.Q.get()
                    if response[0]:
                        self.VER_NO = response[1]
                        self.VER_LINK = response[2]
                        if self.VER_NO is None:
                            messagebox.showerror(master=self, title='UNIPASS ::', message='새로운 버전을 준비 중 입니다.\n잠시 후에 다시 시도해주세요.')
                            self.destroy()
                        else:
                            self.check_version()
                    else:
                        if 'OperationalError' in str(response[1][1]):
                            messagebox.showerror(master=self, title='UNIPASS ::', message='인터넷 연결을 확인해주세요.')
                        self.record_log(response[1])
                        self.destroy()
            else:
                self.CT = CheckVersion(_q=self.Q)
                self.CT.setDaemon(True)
                self.CT.start()
                self.after(100, self.request_version_info, True)
        except:
            self.record_log(sys.exc_info())
            self.destroy()

    @staticmethod
    def record_log(_ei):
        try:
            _dbm = DataBaseManager()
            _dbm.execute_query('''
            INSERT INTO error_report(manager_index, error_code, error_class, error_description, error_line)
            VALUES(0, 'Launcher', '{}', '{}', {});
            '''.format(
                str(_ei[0]).replace('\"', '').replace('\'', ''),
                str(_ei[1]).replace('\"', '').replace('\'', ''),
                int(_ei[2].tb_lineno))
            )
        except:
            with open('./{}.txt'.format(datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d%H%M%S")),
                      'a') as txt:
                txt.write('Launcher' + '\n')
                txt.write(str(_ei[0]) + '\n')
                txt.write(str(_ei[1]) + '\n')
                txt.write(str(_ei[2].tb_lineno) + '\n')
                txt.write(
                    '====================================================================================================\n')
                txt.write('Launcher' + '\n')
                txt.write(str(sys.exc_info()[0]) + '\n')
                txt.write(str(sys.exc_info()[1]) + '\n')
                txt.write(str(sys.exc_info()[2].tb_lineno) + '\n')
                txt.write(
                    '====================================================================================================\n')


class CheckRunStatus(threading.Thread):
    def __init__(self, _q, _ep):
        try:
            super().__init__()
            self._q, self._ep = _q, _ep
        except:
            self._q.put([False, sys.exc_info()])

    def run(self):
        try:
            _run_status = False
            _ver_list = [_exe for _exe in os.listdir(self._ep) if '.exe' in str(_exe)]
            for proc in psutil.process_iter():
                if proc.name() in _ver_list:
                    _run_status = True
                    break
            self._q.put([True, _run_status])
        except:
            self._q.put([False, sys.exc_info()])


class CheckVersion(threading.Thread):
    def __init__(self, _q):
        try:
            super().__init__()
            self._q = _q
            self.DBM = DataBaseManager()
        except:
            self._q.put([False, sys.exc_info()])

    def run(self):
        try:
            ver_query = 'SELECT ver_con_no, ver_con_link FROM ver_con WHERE ver_con_flag = 1 ORDER BY ver_con_index DESC LIMIT 1;'
            ver_result = self.DBM.execute_query(ver_query)
            if len(ver_result) > 0:
                self._q.put([True, ver_result[0][0], ver_result[0][1]])
            else:
                self._q.put([True, None, None])
        except:
            self._q.put([False, sys.exc_info()])
        finally:
            self.DBM.disconnecting()


class Download(threading.Thread):
    def __init__(self, _q, _ep, _vn, _vl):
        try:
            super().__init__()
            self._q, self._ep, self._vn, self._vl = _q, _ep, _vn, _vl
            self.DBM = DataBaseManager()
        except:
            self._q.put([False, sys.exc_info()])

    def run(self):
        try:
            url_list = {
                'ICON': {
                    'PATH': '/'.join(str(self._ep).split('/')[:-1]) + '/icon.ico',
                    'URL': r''.format(r'Key')
                },
                'ZIP': {
                    'PATH': self._ep,
                    'URL': r''.format(self._vl)
                }
            }
            for _k, _v in url_list.items():
                if _k == 'ICON':
                    if os.path.isfile(_v['PATH']):
                        continue
                session = requests.Session()
                response = session.get(_v['URL'], stream=True)
                if b'html' in response.content:
                    _url = str(response.text).split('downloadForm\" action=\"')[-1].split('\" method=\"')[0].replace('amp;', '')
                    response = session.get(_url, stream=True)
                token = None
                for key, value in response.cookies.items():
                    print(key, value)
                    if key.startswith('download_warning'):
                        token = value
                if token:
                    params = {'confirm': token}
                    response = session.get(_v['URL'], params=params, stream=True)
                _cs = 32768
                with open(_v['PATH'], "wb") as f:
                    for chunk in response.iter_content(_cs):
                        if chunk:
                            f.write(chunk)
            self._q.put([True])
        except:
            self._q.put([False, sys.exc_info()])
        finally:
            self.DBM.disconnecting()


class ZipExtract(threading.Thread):
    def __init__(self, _q, _ep, _vn):
        try:
            super().__init__()
            self._q, self._ep, self._vn = _q, _ep, _vn
            self.DBM = DataBaseManager()
        except:
            self._q.put([False, sys.exc_info()])

    def run(self):
        try:
            _zf = zipfile.ZipFile(self._ep)
            _zf.extractall('/'.join(str(self._ep).split('/')[0:-1]))
            self._q.put([True])
        except:
            self._q.put([False, sys.exc_info()])
        finally:
            self.DBM.disconnecting()


class DataBaseManager:
    conn: pymysql.connect

    info: dict

    def __init__(self):
        self.info = {}
        self.connecting(self.info)

    def execute_query(self, sql: str):
        _sql = sql.replace("\n", " ").strip()
        _retry_count = 1
        _execute_status = True

        while _execute_status:
            _execute_result = True
            _response = None
            try:
                with self.conn.cursor() as _conn:
                    _conn.execute(_sql)
                    self.conn.commit()

                    if _sql.lower().strip().startswith("insert", 0):
                        _response = _conn.lastrowid
                    else:
                        _response = _conn.fetchall()
            except:
                _execute_result = False

                if _retry_count < 10:
                    self.conn.ping(reconnect=True)
                    _retry_count += 1
                else:
                    _execute_status = False
            finally:
                if _execute_result:
                    _execute_status = False
                    return _response

    def connecting(self, info):
        self.info = info
        self.conn = pymysql.connect(
            host=self.info['HOST'],
            port=self.info['PORT'],
            user=self.info['USER'],
            password=self.info['PW'],
            database=self.info['DB']
        )

    def disconnecting(self):
        conn = self.conn
        conn.close()


class DataBaseConnectException(BaseException):
    def __init__(self):
        pass


if __name__ == '__main__':
    L: Launcher = Launcher()
    try:
        if not shell.IsUserAnAdmin():
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
            shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
            sys.exit(0)
        L.check_env_root()
        L.mainloop()
    except SystemExit:
        pass
    except:
        L.record_log(sys.exc_info())
