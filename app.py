import os
import sys
import queue
import psutil
import tkinter
import datetime
import pywintypes
from tkinter import messagebox
from win32com.shell import shell
from View import Header, Navigation
from View.view_login import LoginBody
from View.Resource.style import APP_FONT
from View.view_bl_list import BlListBody
from View.view_common import Popup, Shorts
from View.view_vehicle_list import VehicleListBody
from View.view_user_setting import UserSettingBody
from View.view_admin_setting import AdminSettingBody
from View.view_vehicle_lookup import VehicleLookUpBody
from View.view_lookup_bl_history import LookUpBlHistoryBody
from Model.model_common import ErrorReportToDB, ErrorReportToLocal
from View.view_lookup_vehicle_history import LookUpVehicleHistoryBody
from View.Resource.dictionary import VEHICLE_HEAD_CATEGORY, BL_HEAD_CATEGORY, COMMON_HEAD_CATEGORY


class App(tkinter.Tk):
    def __init__(self):
        try:
            self._code = 'App'
            self.VER = '3.1.0'
            self.COPY_RIGHT = 'ABLE GLOBAL LOGISTICS.'
            self.PATH = {}
            self.VEHICLE_HEAD_CATEGORY = VEHICLE_HEAD_CATEGORY
            self.BL_HEAD_CATEGORY = BL_HEAD_CATEGORY
            self.COMMON_HEAD_CATEGORY = COMMON_HEAD_CATEGORY
            for proc in psutil.process_iter():
                if proc.name() == 'UNIPASS For AGL.exe':
                    proc.kill()
                    break
            sys.path.append(os.getcwd())
            for k, v in enumerate(os.environ):
                if v == 'PROGRAMFILES' or v == 'PROGRAMFILES(X86)':
                    self.PATH['MASTER'] = os.environ.get(v).replace('\\', '/')
                    break
            self.PATH['MASTER'] = self.PATH['MASTER'] + '/UNIPASS-AGL'
            self.PATH['FA'] = self.PATH['MASTER'] + '/Config/account.ini'
            self.PATH['IC'] = self.PATH['MASTER'] + '/Icon.ico'
            self.HEADER, self.NAVI, self.BODY = None, None, None
            self.WH, self.WW, self.FH, self.FW = 200, 300, 0, 0
            self.APP_QUEUE, self.THREAD_QUEUE = queue.Queue(), queue.Queue()
            self.CURRENT_THREAD_INFO = {}
            super().__init__()
            self.geometry('{}x{}+{}+{}'.format(
                self.WW,
                self.WH,
                int(self.winfo_screenwidth() / 2) - int(self.WW / 2),
                int(self.winfo_screenheight() / 2) - int(self.WH / 2)
            ))
            self.resizable(False, False)
            self.title('UNIPASS :: {}'.format(self.VER))
            self.iconbitmap(self.PATH['IC'])
            self.APP_DATA = {
                'MI': tkinter.StringVar(self, name='MI', value='0'),
                'MID': tkinter.StringVar(self, name='MID', value=''),
                'MN': tkinter.StringVar(self, name='MN', value=''),
                'MP': tkinter.StringVar(self, name='MP', value=''),
                'MA': tkinter.StringVar(self, name='MA', value='0')
            }
            self.option_add('*TCombobox*Listbox.font', APP_FONT)
            self.POPUP = Popup(self)
            self.SHORTS = Shorts(self)
            self.BODY = LoginBody(self)
            self.BODY.initial()
        except:
            self.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Init',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
        finally:
            self.mainloop()

    def get_current_geometry(self):
        _w = int(self.geometry().split('x')[0])
        _h = int(self.geometry().split('x')[1].split('+')[0])
        _x = int(str(self.geometry()).split('+')[1:][0])
        _y = int(str(self.geometry()).split('+')[1:][1])
        return _w, _h, _x, _y

    def clear_queues(self):
        while not self.APP_QUEUE.empty():
            try:
                self.APP_QUEUE.get()
            except queue.Empty():
                break
        while not self.THREAD_QUEUE.empty():
            try:
                self.THREAD_QUEUE.get()
            except queue.Empty():
                break

    def request_error_report_db(self, _is_run: bool = False, _error_info: [dict, None] = None):
        try:
            if _is_run:
                if self.THREAD_QUEUE.empty():
                    self.after(100, self.request_error_report_db, True)
                else:
                    response = self.THREAD_QUEUE.get()
                    self.POPUP.destroy()
                    if response['SUCCESS']:
                        messagebox.showerror('UNIPASS :: ', '발생한 오류가 기록되었습니다.\n담당자에게 DB 오류 로그의 확인을 요청해주세요.')
                        if self.CURRENT_THREAD_INFO['THREAD'].isAlive():
                            self.CURRENT_THREAD_INFO['THREAD'].join()
                        self.CURRENT_THREAD_INFO.clear()
                    else:
                        self.request_error_report_local(
                            _origin_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'],
                            _second_error_info=response['DATA']
                        )
            else:
                _error_info['ECD'] = self.VER + ':' + _error_info['ECD']
                self.CURRENT_THREAD_INFO = {
                    'THREAD': ErrorReportToDB(
                        _thread_queue=self.THREAD_QUEUE,
                        _manager_index=int(self.getvar('MI')),
                        _error_info=_error_info
                    ),
                    'ERROR_INFO': _error_info
                }
                self.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                self.CURRENT_THREAD_INFO['THREAD'].start()
                self.after(100, self.request_error_report_db, True)
                self.POPUP.loading('발생한 오류를 DB에 기록 중 입니다 ...')
        except:
            self.POPUP.destroy()
            self.request_error_report_local(
                _origin_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'],
                _second_error_info={
                    'ECD': 'RequestErrorReportDB_Run',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [self.CURRENT_THREAD_INFO['ERROR_INFO']]
                }
            )

    def request_error_report_local(self, _is_run: bool = False, _origin_error_info: [dict, None] = None, _second_error_info: [dict, None] = None):
        try:
            if _is_run:
                if self.THREAD_QUEUE.empty():
                    self.after(100, self.request_error_report_local, True)
                else:
                    _response = self.THREAD_QUEUE.get()
                    self.POPUP.destroy()
                    if _response['SUCCESS']:
                        messagebox.showerror('UNIPASS :: ', '발생한 오류가 기록되었습니다.\n담당자에게 아래 경로의 파일을 보내주시거나\n내용을 복사해서 보내주세요.\n\n{}'.format(_response['DATA']))
                        if self.CURRENT_THREAD_INFO['THREAD'].isAlive():
                            self.CURRENT_THREAD_INFO['THREAD'].join()
                        self.CURRENT_THREAD_INFO.clear()
                    else:
                        self.request_error_report_user(
                            _origin_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'][0],
                            _second_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'][1],
                            _third_error_info=_response['DATA']
                        )
            else:
                if self.CURRENT_THREAD_INFO['THREAD'].isAlive():
                    self.CURRENT_THREAD_INFO['THREAD'].join()
                _second_error_info['ECD'] = self.VER + ':' + _second_error_info['ECD']
                self.CURRENT_THREAD_INFO = {
                    'THREAD': ErrorReportToLocal(
                        _thread_queue=self.THREAD_QUEUE,
                        _manager_index=int(self.getvar('MI')),
                        _origin_error_info=_origin_error_info,
                        _second_error_info=_second_error_info,
                        _local_master_path=self.PATH['MASTER']
                    ),
                    'ERROR_INFO': [
                        _origin_error_info, _second_error_info
                    ]
                }
                self.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                self.CURRENT_THREAD_INFO['THREAD'].start()
                self.after(100, self.request_error_report_local, True)
                self.POPUP.loading('발생한 오류를 PC에 기록 중 입니다 ...')
        except:
            self.POPUP.destroy()
            self.request_error_report_user(
                _origin_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'][0],
                _second_error_info=self.CURRENT_THREAD_INFO['ERROR_INFO'][1],
                _third_error_info={
                    'ECD': 'RequestErrorReportLocal_Run',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': self.CURRENT_THREAD_INFO['ERROR_INFO']
                }
            )

    def request_error_report_user(self, _origin_error_info: dict, _second_error_info: dict, _third_error_info: dict):
        _third_error_info['ECD'] = self.VER + ':' + _third_error_info['ECD']
        self.CURRENT_THREAD_INFO['THREAD'].join()
        self.CURRENT_THREAD_INFO.clear()
        _error_report_lines = 'ERROR OCCURRED USER : ' + str(self.getvar('MI')) + '\n'
        _error_report_lines += 'ERROR OCCURRED TIME : ' + str(datetime.datetime.now()) + '\n'
        _error_report_lines += 'ERROR OCCURRED CODE : ' + _origin_error_info['ECD'] + '\n'
        _error_report_lines += 'ERROR OCCURRED CLSS : ' + str(_origin_error_info['CLS']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DESC : ' + str(_origin_error_info['DES']) + '\n'
        _error_report_lines += 'ERROR OCCURRED LNUM : ' + str(_origin_error_info['LNO']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DATA : ' + str(_origin_error_info['DTA']) + '\n'
        _error_report_lines += '====================================================================================================\n'
        _error_report_lines += 'ERROR OCCURRED USER : ' + str(self.getvar('MI')) + '\n'
        _error_report_lines += 'ERROR OCCURRED TIME : ' + str(datetime.datetime.now()) + '\n'
        _error_report_lines += 'ERROR OCCURRED CODE : ' + _second_error_info['ECD'] + '\n'
        _error_report_lines += 'ERROR OCCURRED CLSS : ' + str(_second_error_info['CLS']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DESC : ' + str(_second_error_info['DES']) + '\n'
        _error_report_lines += 'ERROR OCCURRED LNUM : ' + str(_second_error_info['LNO']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DATA : ' + str(_second_error_info['DTA']) + '\n'
        _error_report_lines += '====================================================================================================\n'
        _error_report_lines += 'ERROR OCCURRED USER : ' + str(self.getvar('MI')) + '\n'
        _error_report_lines += 'ERROR OCCURRED TIME : ' + str(datetime.datetime.now()) + '\n'
        _error_report_lines += 'ERROR OCCURRED CODE : ' + _third_error_info['ECD'] + '\n'
        _error_report_lines += 'ERROR OCCURRED CLSS : ' + str(_third_error_info['CLS']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DESC : ' + str(_third_error_info['DES']) + '\n'
        _error_report_lines += 'ERROR OCCURRED LNUM : ' + str(_third_error_info['LNO']) + '\n'
        _error_report_lines += 'ERROR OCCURRED DATA : ' + str(_third_error_info['DTA']) + '\n'
        _error_report_lines += '====================================================================================================\n'
        _error_report_to_user = tkinter.Tk()
        _error_report_to_user.title('UNIPASS ::')
        _error_report_to_user.lift()
        _error_report_to_user.grab_set()
        _error_report_to_user.iconbitmap(self.PATH['IC'])
        _error_report_to_user.state('zoomed')
        _error_desc = tkinter.Label(_error_report_to_user, text='오류가 발생하였습니다.\n\n발생한 오류를 기록할 수 없는 상태입니다.\n\n오류 해결을 위해 반드시 아래 내용을 담당자에게 전달한 후 문의해주세요.', justify=tkinter.LEFT, font=20)
        _error_desc.pack(anchor=tkinter.W, padx=10, pady=10)
        _error_text = tkinter.Text(_error_report_to_user, padx=10, pady=10)
        _error_text.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
        _error_report_line_list = _error_report_lines.split('\n')
        for _error_report_line_no in range(1, len(_error_report_line_list) + 1):
            _error_text.insert('{}.0'.format(_error_report_line_no), _error_report_line_list[_error_report_line_no - 1] + '\n')
        _error_text.configure(state=tkinter.DISABLED)
        _error_report_to_user.mainloop()

    def body_clear(self):
        self.BODY.destroy_widget()
        self.BODY.destroy()

    def body_set_main(self):
        self.body_clear()
        self.resizable(True, True)
        self.minsize(1280, 720)
        self.WW, self.WH = 1280, 720
        self.geometry('{}x{}+{}+{}'.format(
            self.WW,
            self.WH,
            int(self.winfo_screenwidth() / 2) - int(self.WW / 2),
            int(self.winfo_screenheight() / 2) - int(self.WH / 2)
        ))
        self.HEADER = Header(self)
        self.HEADER.set_man_name(self.getvar('MN'))
        self.NAVI = Navigation(self)
        self.NAVI.insert_button('navi_button_vehicle_lookup', '차량 조회', self.body_set_vehicle_lookup)
        self.NAVI.insert_button('navi_button_vehicle_list', '차량 목록', self.body_set_vehicle_list)
        self.NAVI.insert_button('navi_button_lookup_history', '차량 조회 내역', self.body_set_lookup_history)
        self.NAVI.insert_divide()
        self.NAVI.insert_button('navi_button_bl_list', 'B/L 목록', self.body_set_bl_list)
        self.NAVI.insert_button('navi_button_bl_history', 'B/L 조회 내역', self.body_set_bl_history)
        self.NAVI.insert_divide()
        self.NAVI.insert_button('navi_button_user_setting', '사용자 설정', self.body_set_user_setting)
        if int(self.getvar('MA')) == 1:
            self.NAVI.insert_button('navi_button_admin_setting', '계정 관리', self.body_set_admin_setting)
        self.NAVI.insert_button('navi_button_exit', '종료', self.destroy, tkinter.BOTTOM)
        self.body_set_vehicle_lookup()

    def body_set_vehicle_lookup(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_vehicle_lookup')
            self.body_clear()
            self.HEADER.set_title('차량 조회')
            self.BODY = VehicleLookUpBody(self)
            self.BODY.initial()

    def body_set_vehicle_list(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_vehicle_list')
            self.body_clear()
            self.HEADER.set_title('차량 목록')
            self.BODY = VehicleListBody(self)
            self.BODY.initial()

    def body_set_lookup_history(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_lookup_history')
            self.body_clear()
            self.HEADER.set_title('차량 조회 내역')
            self.BODY = LookUpVehicleHistoryBody(self)
            self.BODY.initial()

    def body_set_bl_list(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_bl_list')
            self.body_clear()
            self.HEADER.set_title('B/L 목록')
            self.BODY = BlListBody(self)
            self.BODY.initial()

    def body_set_bl_history(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_bl_history')
            self.body_clear()
            self.HEADER.set_title('B/L 조회 내역')
            self.BODY = LookUpBlHistoryBody(self)
            self.BODY.initial()

    def body_set_user_setting(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_user_setting')
            self.body_clear()
            self.HEADER.set_title('사용자 설정')
            self.BODY = UserSettingBody(self)
            self.BODY.initial()

    def body_set_admin_setting(self):
        if not self.POPUP.winfo_exists():
            self.NAVI.disable_button('navi_button_admin_setting')
            self.body_clear()
            self.HEADER.set_title('계정 관리')
            self.BODY = AdminSettingBody(self)
            self.BODY.initial()

    def destroy(self) -> None:
        if len(self.CURRENT_THREAD_INFO) == 0:
            self.body_clear()
            super().destroy()


if __name__ == '__main__':
    try:
        if not shell.IsUserAnAdmin():
            script = os.path.abspath(sys.argv[0])
            params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
            shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
            sys.exit(0)
        APP = App()
    except SystemExit:
        pass
    except pywintypes.error:
        messagebox.showerror('UNIPASS ::', '관리자 권한 습득에 실패하였습니다.')
