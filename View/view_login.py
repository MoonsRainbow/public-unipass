import sys
import tkinter
from tkinter import messagebox
from View import Body, TitleEntry
from Model.model_login import (GetLocalData, SetLocalData, GetAuthority)


class LoginBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewLogin_LoginBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA = {
                'LB_VAR_SAVE': tkinter.BooleanVar(master=self.APP, value=False),
                'IS_AUTHORITY': tkinter.BooleanVar(master=self.APP, value=False)
            }
            self.WIDGETS.clear()
            self.WIDGETS = {
                'LB_TITLE_ENTRY_ID': TitleEntry(self, 'ID'),
                'LB_TITLE_ENTRY_PW': TitleEntry(self, 'PW', True),
                'LB_CHECK_BUTTON_SAVE': tkinter.Checkbutton(self, text='ID/PW 저장', variable=self.DATA['LB_VAR_SAVE']),
                'LB_BUTTON_LOGIN': tkinter.Button(self, text='로 그 인', overrelief=tkinter.SOLID, command=self.event_button_click_login, repeatdelay=1000, repeatinterval=100, height=3)
            }
            self.request_get_local_data()
            self.WIDGETS['LB_TITLE_ENTRY_ID'].place(relwidth=1, height=44)
            self.WIDGETS['LB_TITLE_ENTRY_PW'].place(y=54, relwidth=1, height=44)
            self.WIDGETS['LB_CHECK_BUTTON_SAVE'].place(x=180, y=108, width=100, height=18)
            self.WIDGETS['LB_BUTTON_LOGIN'].place(y=136, relwidth=1, height=44)
            self.pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
            self.APP.bind('<Return>', self.event_button_click_login)
        except:
            self.APP.request_error_report_db(
                _error_info= {
                    'ECD': self._code + '_Initial',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            self.APP.destroy()

    def event_button_click_login(self, _event=None):
        try:
            if self.WIDGETS['LB_TITLE_ENTRY_ID'].get_value() != '' and self.WIDGETS['LB_TITLE_ENTRY_PW'].get_value() != '':
                self.request_get_authority()
                if self.DATA['IS_AUTHORITY'].get():
                    self.request_set_local_data()
                    self.APP.body_set_main()
            else:
                messagebox.showinfo('UNIPASS :: ', '아이디와 비밀번호가 모두 입력되지 않았습니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickLogin',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '로그인 중 오류가 발생하였습니다.')
            self.APP.destroy()

    def request_get_local_data(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_local_data, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.WIDGETS['LB_TITLE_ENTRY_ID'].set_value(_response['DATA']['ID'])
                        self.WIDGETS['LB_TITLE_ENTRY_PW'].set_value(_response['DATA']['PW'])
                        if _response['DATA']['ID'] != '' and _response['DATA']['PW'] != '':
                            self.DATA['LB_VAR_SAVE'].set(True)
                        else:
                            self.DATA['LB_VAR_SAVE'].set(False)
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        self.WIDGETS['LB_TITLE_ENTRY_ID'].set_value('')
                        self.WIDGETS['LB_TITLE_ENTRY_PW'].set_value('')
                        messagebox.showerror('UNIPASS :: ', '저장된 계정 정보를 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                    'THREAD': GetLocalData(
                        _thread_queue=self.APP.THREAD_QUEUE,
                        _local_file_path=self.APP.PATH['FA'],
                        _key=self.APP.COPY_RIGHT
                    )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_local_data, True)
                    self.APP.POPUP.loading('저장된 계정 정보를 불러오는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetLocalData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            self.WIDGETS['LB_TITLE_ENTRY_ID'].set_value('')
            self.WIDGETS['LB_TITLE_ENTRY_PW'].set_value('')
            messagebox.showerror('UNIPASS :: ', '저장된 계정 정보를 불러오는 중 오류가 발생하였습니다.')

    def request_set_local_data(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_set_local_data, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if not _response['SUCCESS']:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '계정 정보를 저장하는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': SetLocalData(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _local_file_path=self.APP.PATH['FA'],
                            _key=self.APP.COPY_RIGHT,
                            _login_info={
                                'ID': self.WIDGETS['LB_TITLE_ENTRY_ID'].get_value() if self.DATA['LB_VAR_SAVE'].get() else 'Deleted',
                                'PW': self.WIDGETS['LB_TITLE_ENTRY_PW'].get_value() if self.DATA['LB_VAR_SAVE'].get() else 'Deleted'
                            }
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_set_local_data, True)
                    self.APP.POPUP.loading('계정 정보를 저장하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestSetLocalData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 정보를 저장하는 중 오류가 발생하였습니다.')

    def request_get_authority(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_authority, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        _response['DATA'] = _response['DATA'][0]
                        if _response['DATA'][0] == 0:
                            self.DATA['IS_AUTHORITY'].set(False)
                            messagebox.showinfo(title='알림 :: ', message='로그인에 실패하였습니다.\n계정 정보를 다시 확인해주세요.')
                        else:
                            self.DATA['IS_AUTHORITY'].set(True)
                            self.APP.setvar('MI', _response['DATA'][1])
                            self.APP.setvar('MID', self.WIDGETS['LB_TITLE_ENTRY_ID'].get_value())
                            self.APP.setvar('MN', _response['DATA'][2])
                            self.APP.setvar('MP', _response['DATA'][3])
                            self.APP.setvar('MA', _response['DATA'][4])
                    else:
                        self.DATA['IS_AUTHORITY'].set(False)
                        if 'OperationalError' in str(_response['DATA']['CLS'][0]):
                            messagebox.showerror(title='UNIPASS ::', message='인터넷 연결을 확인해주세요.')
                        else:
                            self.APP.request_error_report_db(_error_info=_response['DATA'])
                            self.APP.destroy()
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetAuthority(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _login_info={
                                'ID': self.WIDGETS['LB_TITLE_ENTRY_ID'].get_value(),
                                'PW': self.WIDGETS['LB_TITLE_ENTRY_PW'].get_value()
                            }
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_authority, True)
                    self.APP.POPUP.loading('로그인 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetAuthority',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '로그인 중 오류가 발생하였습니다.')
            self.APP.destroy()
