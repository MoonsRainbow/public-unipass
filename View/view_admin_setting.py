import sys
import queue
import tkinter
from View import Body, Table
from tkinter import messagebox
from Model.model_admin_setting import GetManagerList, AddManager, DeleteManager, ChangeManagerPassword


class AdminSettingBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewAdminSetting_AdminSettingBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA.clear()
            self.WIDGETS.clear()
            self.WIDGETS['B_ADD_MANAGER'] = tkinter.Button(self, text='계정 추가', overrelief=tkinter.SOLID, command=self.event_button_click_manager_add, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['TABLE'] = Table(self.APP, self)
            self.WIDGETS['B_ADD_MANAGER'].place(relx=0.88, relwidth=0.12, height=56)
            self.WIDGETS['TABLE'].set_location_y(86)
            self.WIDGETS['TABLE'].LOCATION.place(y=66, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 86)
            self.view_reset()
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
            self.WIDGETS['TABLE'].set_table_resizing()
            self.WIDGETS['TABLE'].set_table_paging()
            self.WIDGETS['TABLE'].set_table_celling()
            self.WIDGETS['TABLE'].set_head_sort_command()
            self.WIDGETS['TABLE'].set_cell_double_click_command()
            self.WIDGETS['TABLE'].bind('<Double-1>', self.event_table_double_click)
            self.WIDGETS['TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.COMMON_HEAD_CATEGORY.items() if 'ASML' in _value['TABLE']},
                _is_change=True
            )
            self.request_get_manager_list()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ResetView',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '화면 초기화 중 오류가 발생하였습니다.')

    def event_button_click_manager_add(self):
        try:
            self.APP.clear_queues()
            self.APP.POPUP.add_manager(self.DATA['ADD_MANAGER_INFO'])
            if self.APP.APP_QUEUE.empty():
                self.DATA.clear()
            else:
                self.DATA['ADD_MANAGER_INFO'] = self.APP.APP_QUEUE.get()
                self.APP.clear_queues()
                self.request_add_manager()
                if self.APP.APP_QUEUE.empty():
                    self.DATA.clear()
                else:
                    if self.APP.APP_QUEUE.get():
                        self.view_reset()
                    else:
                        messagebox.showwarning('UNIPASS ::', '중복된 아이디가 있습니다.\n다시 입력해주세요.')
                        self.APP.after(0, self.event_button_click_manager_add)
        except KeyError:
            self.DATA['ADD_MANAGER_INFO'] = None
            self.APP.after(0, self.event_button_click_manager_add)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickManagerAdd',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 추가 중 오류가 발생하였습니다.')

    def event_table_double_click(self, _event=None):
        try:
            if not self.APP.POPUP.winfo_exists():
                self.DATA['SELECTED_MANAGER_INDEX'] = list(self.WIDGETS['TABLE'].item(self.WIDGETS['TABLE'].selection()[0]).values())[2][0]
                self.DATA['SELECTED_MANAGER_ID'] = list(self.WIDGETS['TABLE'].item(self.WIDGETS['TABLE'].selection()[0]).values())[2][1]
                self.WIDGETS['TABLE'].selection_remove(self.WIDGETS['TABLE'].selection()[0])
                self.APP.clear_queues()
                self.APP.POPUP.manager_menu(self.DATA['SELECTED_MANAGER_ID'])
                if self.APP.APP_QUEUE.empty():
                    self.DATA.clear()
                else:
                    _selected_manager_menu = self.APP.APP_QUEUE.get()
                    if _selected_manager_menu == 1:
                        self.APP.clear_queues()
                        self.APP.POPUP.change_manager_password()
                        if not self.APP.APP_QUEUE.empty():
                            self.DATA['CHANGE_MANAGER_PASSWORD'] = self.APP.APP_QUEUE.get()
                            self.request_change_manager_password()
                    else:
                        if messagebox.askyesno('UNIPASS ::', '\'{}\' 계정을 삭제할까요?'.format(self.DATA['SELECTED_MANAGER_ID'])):
                            self.request_delete_manager()
                            self.DATA.clear()
        except IndexError:
            if not self.APP.POPUP.winfo_exists():
                self.DATA.clear()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventTableDoubleClick',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 관리 메뉴 실행 중 오류가 발생하였습니다.')

    def request_get_manager_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.after(100, self.request_get_manager_list, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.WIDGETS['TABLE'].delete_row()
                        self.WIDGETS['TABLE'].insert_row(_response['DATA'])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '계정 목록을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetManagerList(
                            _thread_queue=self.APP.THREAD_QUEUE
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_manager_list, True)
                    self.APP.POPUP.loading('계정 목록을 불러오는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetManagerList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 목록을 불러오는 중 오류가 발생하였습니다.')

    def request_add_manager(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_add_manager, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.APP.APP_QUEUE.put(_response['DATA'])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '계정 추가 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': AddManager(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_info=self.DATA['ADD_MANAGER_INFO']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_add_manager, True)
                    self.APP.POPUP.loading('계정을 추가하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestAddManager',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 추가 중 오류가 발생하였습니다.')

    def request_delete_manager(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_delete_manager, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.view_reset()
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '계정 삭제 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': DeleteManager(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_index=self.DATA['SELECTED_MANAGER_INDEX']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_delete_manager, True)
                    self.APP.POPUP.loading('계정을 삭제하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestDeleteManager',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 삭제 중 오류가 발생하였습니다.')

    def request_change_manager_password(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_change_manager_password, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if not _response['SUCCESS']:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '계정 비밀번호 변경 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': ChangeManagerPassword(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_index=self.DATA['SELECTED_MANAGER_INDEX'],
                            _manager_password=self.DATA['CHANGE_MANAGER_PASSWORD']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_change_manager_password, True)
                    self.APP.POPUP.loading('계정 비밀번호를 변경하는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestChangeManagerPassword',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '계정 비밀번호 변경 중 오류가 발생하였습니다.')
