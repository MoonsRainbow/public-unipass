import sys
import tkinter
from tkinter import ttk
from tkinter import messagebox, filedialog
from View import Body, LabelEnvDownloadPath
from Model.model_user_setting import SetDownloadPath, GetExcelColTitleList, AddExcelTitle, DeleteExcelTitle


class UserSettingBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewUserSetting_UserSettingBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA.clear()
            self.DATA['CATEGORY_CET'] = {}
            self.DATA['CATEGORY_BET'] = {}
            self.DATA['TITLE_INFO'] = {}
            self.WIDGETS.clear()
            self.WIDGETS['LE_DOWNLOAD_PATH'] = LabelEnvDownloadPath(self, '다운로드 위치 설정', '엑셀이 다운로드될 기본 위치를 설정합니다.', self.APP.APP_DATA['MP'], '설정', self.event_button_click_download_path_setting)
            self.WIDGETS['LE_DOWNLOAD_PATH'].place(relwidth=1, height=100)
            self.WIDGETS['LE_CET'] = tkinter.LabelFrame(self, text='엑셀 열 제목 설정 (차대번호)', padx=5, pady=5, foreground='gray')
            self.WIDGETS['B_CET_ADD'] = tkinter.Button(self.WIDGETS['LE_CET'], text='추가', overrelief=tkinter.SOLID, command=self.event_button_click_cet_add, repeatdelay=1000, repeatinterval=100, height=2)
            self.WIDGETS['CET_SY'] = tkinter.Scrollbar(self.WIDGETS['LE_CET'])
            self.WIDGETS['CET_SX'] = tkinter.Scrollbar(self.WIDGETS['LE_CET'], orient=tkinter.HORIZONTAL)
            self.WIDGETS['CET_T'] = ttk.Treeview(self.WIDGETS['LE_CET'], show='headings', xscrollcommand=self.WIDGETS['CET_SX'].set, yscrollcommand=self.WIDGETS['CET_SY'].set)
            self.WIDGETS['CET_SY'].config(command=self.WIDGETS['CET_T'].yview)
            self.WIDGETS['CET_SX'].config(command=self.WIDGETS['CET_T'].xview)
            self.WIDGETS['CET_T']['columns'] = ['No', 'ColumnTitle']
            self.WIDGETS['CET_T'].column('No', anchor=tkinter.CENTER, width=40, stretch=tkinter.NO)
            self.WIDGETS['CET_T'].heading('No', text='No')
            self.WIDGETS['CET_T'].column('ColumnTitle', anchor=tkinter.W, width=230, stretch=tkinter.NO)
            self.WIDGETS['CET_T'].heading('ColumnTitle', text='열 제목')
            self.WIDGETS['LE_CET'].place(y=110, relwidth=0.4955, height=self.APP.get_current_geometry()[1] * 0.94 - 130)
            tkinter.Label(self.WIDGETS['LE_CET'], text='엑셀에서 찾을 차대번호 열의 제목을 설정합니다.').pack(side=tkinter.TOP, anchor=tkinter.W)
            self.WIDGETS['CET_SY'].pack(side=tkinter.RIGHT, fill=tkinter.Y)
            self.WIDGETS['CET_SX'].pack(side=tkinter.BOTTOM, fill=tkinter.X)
            self.WIDGETS['CET_T'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE, pady=10)
            self.WIDGETS['B_CET_ADD'].pack(side=tkinter.TOP, fill=tkinter.X)
            self.WIDGETS['CET_T'].bind('<Double-1>', self.event_cet_table_cell_double_click)
            self.WIDGETS['LE_BET'] = tkinter.LabelFrame(self, text='엑셀 열 제목 설정 (B/L 번호)', padx=5, pady=5, foreground='gray')
            self.WIDGETS['B_BET_ADD'] = tkinter.Button(self.WIDGETS['LE_BET'], text='추가', overrelief=tkinter.SOLID, command=self.event_button_click_bet_add, repeatdelay=1000, repeatinterval=100, height=2)
            self.WIDGETS['BET_SY'] = tkinter.Scrollbar(self.WIDGETS['LE_BET'])
            self.WIDGETS['BET_SX'] = tkinter.Scrollbar(self.WIDGETS['LE_BET'], orient=tkinter.HORIZONTAL)
            self.WIDGETS['BET_T'] = ttk.Treeview(self.WIDGETS['LE_BET'], show='headings', xscrollcommand=self.WIDGETS['BET_SX'].set, yscrollcommand=self.WIDGETS['BET_SY'].set)
            self.WIDGETS['BET_SY'].config(command=self.WIDGETS['BET_T'].yview)
            self.WIDGETS['BET_SX'].config(command=self.WIDGETS['BET_T'].xview)
            self.WIDGETS['BET_T']['columns'] = ['No', 'ColumnTitle']
            self.WIDGETS['BET_T'].column('No', anchor=tkinter.CENTER, width=40, stretch=tkinter.NO)
            self.WIDGETS['BET_T'].heading('No', text='No')
            self.WIDGETS['BET_T'].column('ColumnTitle', anchor=tkinter.W, width=230, stretch=tkinter.NO)
            self.WIDGETS['BET_T'].heading('ColumnTitle', text='열 제목')
            self.WIDGETS['LE_BET'].place(y=110, relwidth=0.4955, relx=0.5045, height=self.APP.get_current_geometry()[1] * 0.94 - 130)
            tkinter.Label(self.WIDGETS['LE_BET'], text='엑셀에서 찾을 B/L 번호 열의 제목을 설정합니다.').pack(side=tkinter.TOP, anchor=tkinter.W)
            self.WIDGETS['BET_SY'].pack(side=tkinter.RIGHT, fill=tkinter.Y)
            self.WIDGETS['BET_SX'].pack(side=tkinter.BOTTOM, fill=tkinter.X)
            self.WIDGETS['BET_T'].pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=tkinter.TRUE, pady=10)
            self.WIDGETS['B_BET_ADD'].pack(side=tkinter.TOP, fill=tkinter.X)
            self.WIDGETS['BET_T'].bind('<Double-1>', self.event_bet_table_cell_double_click)
            self.request_get_excel_col_title_list()
            self.body_place()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Initial',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            self.APP.destroy()

    def view_reset(self):
        try:
            self.DATA.clear()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ViewReset',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '화면 초기화 중 오류가 발생하였습니다.')

    def event_button_click_download_path_setting(self):
        try:
            _selected_download_path = filedialog.askdirectory(initialdir=self.getvar('MP') if self.getvar('MP') != '' else './', title='다운로드 위치 설정').replace('\\', '/')
            if _selected_download_path != '' and self.getvar('MP') != _selected_download_path:
                self.DATA['SELECTED_DOWNLOAD_PATH'] = _selected_download_path
                self.request_set_download_path()
                self.setvar('MP', self.DATA['SELECTED_DOWNLOAD_PATH'])
                self.view_reset()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickDownloadPathSetting',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '다운로드 위치 설정 중 오류가 발생하였습니다.')

    def event_button_click_cet_add(self):
        try:
            self.APP.POPUP.add_excel_title()
            if not self.APP.APP_QUEUE.empty():
                self.DATA['TITLE_INFO']['TYPE'] = 1
                self.DATA['TITLE_INFO']['VALUE'] = self.APP.APP_QUEUE.get()
                self.request_add_excel_title()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickCETAdd',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '열 제목 추가 중 오류가 발생하였습니다.')

    def event_cet_table_cell_double_click(self, _event=None):
        if messagebox.askyesno('UNIPASS ::', '열 제목 \n\'{}\' 을 삭제할까요?'.format(
            self.DATA['CATEGORY_CET'][list(self.WIDGETS['CET_T'].item(self.WIDGETS['CET_T'].selection()[0]).values())[2][0]]['TEXT']
        )):
            self.DATA['TITLE_INFO'].__setitem__(
                'INDEX', self.DATA['CATEGORY_CET'][list(self.WIDGETS['CET_T'].item(self.WIDGETS['CET_T'].selection()[0]).values())[2][0]]['INDEX']
            )
            self.request_delete_excel_title()

    def event_button_click_bet_add(self):
        try:
            self.APP.POPUP.add_excel_title()
            if not self.APP.APP_QUEUE.empty():
                self.DATA['TITLE_INFO']['TYPE'] = 2
                self.DATA['TITLE_INFO']['VALUE'] = self.APP.APP_QUEUE.get()
                self.request_add_excel_title()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickBETAdd',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '열 제목 추가 중 오류가 발생하였습니다.')

    def event_bet_table_cell_double_click(self, _event=None):
        if messagebox.askyesno('UNIPASS ::', '열 제목 \n\'{}\' 을 삭제할까요?'.format(
                self.DATA['CATEGORY_BET'][list(self.WIDGETS['BET_T'].item(self.WIDGETS['BET_T'].selection()[0]).values())[2][0]]['TEXT']
        )):
            self.DATA['TITLE_INFO'].__setitem__(
                'INDEX', self.DATA['CATEGORY_BET'][list(self.WIDGETS['BET_T'].item(self.WIDGETS['BET_T'].selection()[0]).values())[2][0]]['INDEX']
            )
            self.request_delete_excel_title()

    def request_get_excel_col_title_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_excel_col_title_list, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        for _row_no in self.WIDGETS['CET_T'].get_children():
                            self.WIDGETS['CET_T'].delete(_row_no)
                        for _row_no in self.WIDGETS['BET_T'].get_children():
                            self.WIDGETS['BET_T'].delete(_row_no)
                        for i in range(len(_response['DATA'][0])):
                            self.DATA['CATEGORY_CET'].__setitem__(i + 1, {'TEXT': _response['DATA'][0][i][1], 'INDEX': _response['DATA'][0][i][0]})
                            self.WIDGETS['CET_T'].insert(parent='', values=[i + 1, _response['DATA'][0][i][1]], index=i)
                        for i in range(len(_response['DATA'][1])):
                            self.DATA['CATEGORY_BET'].__setitem__(i + 1, {'TEXT': _response['DATA'][1][i][1], 'INDEX': _response['DATA'][1][i][0]})
                            self.WIDGETS['BET_T'].insert(parent='', values=[i + 1, _response['DATA'][1][i][1]], index=i)
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '열 제목을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {'THREAD': GetExcelColTitleList(_thread_queue=self.APP.THREAD_QUEUE)}
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_excel_col_title_list, True)
                    self.APP.POPUP.loading('열 제목을 불러오는 중 입니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetExcelColTitleList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '열 제목을 불러오는 중 오류가 발생하였습니다.')

    def request_add_excel_title(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_add_excel_title, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.DATA['TITLE_INFO'].clear()
                        self.request_get_excel_col_title_list()
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '열 제목 추가 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': AddExcelTitle(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _env_type=self.DATA['TITLE_INFO']['TYPE'],
                            _env_value=self.DATA['TITLE_INFO']['VALUE']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_add_excel_title, True)
                    self.APP.POPUP.loading('열 제목을 추가하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestAddExcelTitle',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '열 제목 추가 중 오류가 발생하였습니다.')

    def request_delete_excel_title(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_delete_excel_title, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.DATA['TITLE_INFO'].clear()
                        self.request_get_excel_col_title_list()
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '열 제목 삭제 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': DeleteExcelTitle(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _env_index=self.DATA['TITLE_INFO']['INDEX']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_delete_excel_title, True)
                    self.APP.POPUP.loading('열 제목을 삭제하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestDeleteExcelTitle',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '열 제목 삭제 중 오류가 발생하였습니다.')

    def request_set_download_path(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_set_download_path, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if not _response['SUCCESS']:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '다운로드 위치 저장 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': SetDownloadPath(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_index=self.getvar('MI'),
                            _download_path=self.DATA['SELECTED_DOWNLOAD_PATH']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_set_download_path, True)
                    self.APP.POPUP.loading('다운로드 위치를 저장 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestSetDownloadPath',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '다운로드 위치 저장 중 오류가 발생하였습니다.')
