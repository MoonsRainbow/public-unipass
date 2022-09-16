import sys
import tkinter
import datetime
import win32print
from tkinter import (messagebox, ttk)
from View import LabelText, TitleEntry, LabelCombo


class Shorts(tkinter.Toplevel):
    def __init__(self, _app):
        self._code, self.APP, self._W, self._H = 'Shorts', _app, 0, 0
        super().__init__(self.APP)
        self.geometry('0x0+0+0')
        self.destroy()

    def show(self, _msg: str):
        try:
            if not self.winfo_exists():
                for _widget in [_w for _w in self.children.values()]:
                    _widget.destroy()
                super().__init__(self.APP)
                self._W, self._H = 300, 100
                _w, _h, _x, _y = self.APP.get_current_geometry()
                self.APP.WW, self.APP.WH = _w, _h
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 0.9)
                self.overrideredirect(True)
                self.resizable(False, False)
                LabelText(self, _title='', _var=tkinter.StringVar(master=self, value=_msg)).pack(fill=tkinter.BOTH, expand=tkinter.TRUE, padx=5, pady=5)
                self.lift()
                self.after(500, self.destroy)
            else:
                self.destroy()
                self.show(_msg)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Show',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '알림 팝업 생성 중 오류가 발생하였습니다.')

    def destroy(self) -> None:
        try:
            if self.winfo_exists():
                super().destroy()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Destroy',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '알림 팝업 종료 중 오류가 발생하였습니다.')


class Popup(tkinter.Toplevel):
    def __init__(self, _app):
        self._code, self.APP, self.DATA, self.AFTER, self.WIDGETS, self._W, self._H = 'Popup', _app, {}, {}, {}, 0, 0
        super().__init__(self.APP)
        self.geometry('0x0+0+0')
        self.destroy()

    def initial(self):
        try:
            for _widget in [_w for _w in self.children.values()]:
                _widget.destroy()
            self.DATA.clear()
            self.WIDGETS.clear()
            self.APP.bind('<Enter>', self._after_focusing)
            self.APP.bind('<FocusIn>', self._after_focusing)
            self._W, self._H = 0, 0
            super().__init__(self.APP)
            self.geometry('0x0+0+0')
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
            messagebox.showerror('UNIPASS ::', '팝업 초기화 중 오류가 발생하였습니다.')

    def _after_running(self):
        try:
            if self.DATA['RS'] >= 10:
                self.DATA['RS'] = 0
            else:
                self.DATA['RS'] += 1
            try:
                self.WIDGETS['L_S'].configure(text=' ● ' * self.DATA['RS'] + ' ○ ' * (10 - self.DATA['RS']))
                self.AFTER['AFTER_RS_ID'] = self.after(100, self._after_running)
            except tkinter.TclError:
                pass
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_AfterRunning',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '로딩 업데이트 중 오류가 발생하였습니다.')

    def _after_focusing(self, _event=None):
        try:
            self.lift()
            _last_focus = self.focus_lastfor()
            self.focus_set()
            _last_focus.focus_set()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_AfterFocusing',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '포커싱 업데이트 중 오류가 발생하였습니다.')

    def _after_covering(self, _event=None):
        try:
            self.APP.WW, self.APP.WH = self.APP.get_current_geometry()[:2]
            if self.APP.state() == 'iconic':
                self.geometry('0x0+0+0')
            elif self.APP.state() == 'zoomed':
                _aw, _ah, _ax, _ay = self.APP.get_current_geometry()
                self.geometry('{}x{}+{}+{}'.format(_aw, _ah + 1, 0, 23))
                self.lift()
            else:
                _aw, _ah, _ax, _ay = self.APP.get_current_geometry()
                if _aw > 300:
                    self.geometry('{}x{}+{}+{}'.format(_aw, _ah + 1, _ax + 8, _ay + 30))
                    self.lift()
                else:
                    self.geometry('{}x{}+{}+{}'.format(_aw, _ah + 1, _ax + 3, _ay + 25))
                    self.lift()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_AfterCovering',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '커버링 업데이트 중 오류가 발생하였습니다.')

    def destroy(self, _evnet=None) -> None:
        try:
            if self.winfo_exists():
                for _after_id in self.AFTER.values():
                    self.after_cancel(_after_id)
                self.AFTER.clear()
                self.unbind('<Escape>')
                self.unbind('<Configure>')
                self.APP.unbind('<Enter>')
                self.APP.unbind('<FocusIn>')
                self.APP.unbind_all('<MouseWheel>')
                super().destroy()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Destroy',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '팝업 종료 중 오류가 발생하였습니다.')

    def _set_loading_message(self, _msg: str):
        try:
            if 'MSG' in list(self.DATA.keys()):
                self.DATA['MSG'].set(_msg)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_SetLoadingMessage',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '로딩 메세지 설정 중 오류가 발생하였습니다.')

    def loading(self, _msg: str):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Configure>', self._after_covering)
                self.wm_attributes('-alpha', 0.9)
                self.overrideredirect(True)
                self.resizable(False, False)
                self.title('UNIPASS :: ')
                self.DATA = {
                    'RS': 0,
                    'MSG': tkinter.StringVar(self, value=_msg)
                }
                self.WIDGETS = {
                    'L_S': tkinter.Label(self, text=' ○  ○  ○  ○  ○  ○  ○  ○  ○  ○ '),
                    'L_M': tkinter.Label(self, textvariable=self.DATA['MSG']),
                }
                self.WIDGETS['L_S'].pack(anchor=tkinter.S, expand=tkinter.TRUE)
                self.WIDGETS['L_M'].pack(anchor=tkinter.N, expand=tkinter.TRUE)
                self.lift()
                self._after_running()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_Loading',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '로딩 팝업 생성 중 오류가 발생하였습니다.')

    def _event_table_double_click_sheet_chosen(self, _event):
        try:
            _chosen_index = list(self.WIDGETS['CS_T'].item(self.WIDGETS['CS_T'].selection()[0]).values())[2][0] - 1
            self.APP.THREAD_QUEUE.put(self.DATA['SL'][_chosen_index])
            self.destroy()
        except IndexError:
            pass
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventTableDoubleClickSheetChosen',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '시트 선택 중 오류가 발생하였습니다.')

    def sheet_choice(self, _sheet_list):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 300, 500
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 시트 선택')
                self.iconbitmap(self.APP.PATH['IC'])
                self.DATA = {
                    'SL': _sheet_list,
                }
                self.WIDGETS['CS_F'] = tkinter.Frame(self, padx=5, pady=5)
                self.WIDGETS['CS_SY'] = tkinter.Scrollbar(self.WIDGETS['CS_F'])
                self.WIDGETS['CS_SX'] = tkinter.Scrollbar(self.WIDGETS['CS_F'], orient=tkinter.HORIZONTAL)
                self.WIDGETS['CS_T'] = ttk.Treeview(self.WIDGETS['CS_F'], show='headings', xscrollcommand=self.WIDGETS['CS_SX'].set, yscrollcommand=self.WIDGETS['CS_SY'].set)
                self.WIDGETS['CS_SY'].config(command=self.WIDGETS['CS_T'].yview)
                self.WIDGETS['CS_SX'].config(command=self.WIDGETS['CS_T'].xview)
                self.WIDGETS['CS_T']['columns'] = ['No', 'SheetName']
                self.WIDGETS['CS_T'].column('No', anchor=tkinter.CENTER, width=40, stretch=tkinter.NO)
                self.WIDGETS['CS_T'].heading('No', text='No')
                self.WIDGETS['CS_T'].column('SheetName', anchor=tkinter.W, width=230, stretch=tkinter.NO)
                self.WIDGETS['CS_T'].heading('SheetName', text='시트 이름')
                self.WIDGETS['CS_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                self.WIDGETS['CS_SY'].pack(side=tkinter.RIGHT, fill=tkinter.Y)
                self.WIDGETS['CS_SX'].pack(side=tkinter.BOTTOM, fill=tkinter.X)
                self.WIDGETS['CS_T'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                for i in range(0, len(self.DATA['SL'])):
                    self.WIDGETS['CS_T'].insert(parent='', values=[i + 1, self.DATA['SL'][i]], index=i)
                self.WIDGETS['CS_T'].bind('<Double-1>', lambda event: self._event_table_double_click_sheet_chosen(_event=event))
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_SheetChoice',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '시트 선택 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_lookup_cancel(self):
        try:
            self.WIDGETS['LD_BTN_STOP'].configure(state=tkinter.DISABLED)
            self.DATA['MSG'].set('조회 작업을 중단하는 중 입니다 ...')
            self.APP.APP_QUEUE.put(False)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickLookUpData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '유니패스 조회 취소 처리 중 오류가 발생하였습니다.')

    def lookup_data(self, _msg: str, _var_count_total: [tkinter.StringVar, int]):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Configure>', self._after_covering)
                self.wm_attributes('-alpha', 0.9)
                self.overrideredirect(True)
                self.resizable(False, False)
                self.title('UNIPASS :: ')
                if type(_var_count_total) == int:
                    _var_count_total = tkinter.StringVar(master=self, value=str(_var_count_total))
                self.DATA = {
                    'RS': 0,
                    'TOTAL': _var_count_total,
                    'SUCCESS': tkinter.StringVar(master=self, value='0'),
                    'FAIL': tkinter.StringVar(master=self, value='0'),
                    'MSG': tkinter.StringVar(master=self, value=_msg)
                }
                self.WIDGETS['LD_F'] = tkinter.Frame(self, padx=5, pady=5)
                self.WIDGETS['L_S'] = tkinter.Label(self, text=' ○  ○  ○  ○  ○  ○  ○  ○  ○  ○ ')
                self.WIDGETS['L_M'] = tkinter.Label(self, textvariable=self.DATA['MSG'])
                self.WIDGETS['LD_TOT'] = LabelText(self.WIDGETS['LD_F'], '전체 개수', self.DATA['TOTAL'])
                self.WIDGETS['LD_SUC'] = LabelText(self.WIDGETS['LD_F'], '성공 개수', self.DATA['SUCCESS'])
                self.WIDGETS['LD_FIL'] = LabelText(self.WIDGETS['LD_F'], '실패 개수', self.DATA['FAIL'])
                self.WIDGETS['LD_BTN_STOP'] = tkinter.Button(self, text='취소', overrelief=tkinter.SOLID, command=self._event_button_click_lookup_cancel, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['LD_TOT'].pack(side=tkinter.LEFT)
                self.WIDGETS['LD_SUC'].pack(side=tkinter.LEFT, padx=20)
                self.WIDGETS['LD_FIL'].pack(side=tkinter.LEFT)
                self.WIDGETS['LD_F'].pack(anchor=tkinter.S, expand=tkinter.TRUE, pady=10)
                self.WIDGETS['L_S'].pack(pady=5)
                self.WIDGETS['L_M'].pack(pady=5)
                self.WIDGETS['LD_BTN_STOP'].pack(anchor=tkinter.N, expand=tkinter.TRUE, pady=10)
                self.lift()
                self._after_running()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_LookUpData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '유니패스 조회 팝업 생성 중 오류가 발생하였습니다.')

    def _event_table_cell_click(self, _code, _copy_command, _event=None):
        try:
            if self.WIDGETS[_code].identify_region(_event.x, _event.y) == 'cell':
                _copy_command(
                    _event=_event,
                    _key='BL 번호',
                    _value=list(self.WIDGETS[_code].item(self.WIDGETS[_code].identify_row(_event.y)).values())[2][0]
                )
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventTableCellClick',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_code, _event]
                }
            )
            messagebox.showerror('UNIPASS :: ', '복사에 실패하였습니다.')

    def detail_data(self, _data: dict, _copy_command):
        try:
            if not self.winfo_exists():
                self.initial()
                self.WIDGETS['LDD_F'] = tkinter.Frame(self, padx=10, pady=10)
                self.bind('<Escape>', self.destroy)
                if _data['STATUS_FLAG'] == 2:
                    self._W, self._H = 1000, 800
                    self.WIDGETS['LDD_F_COM'] = tkinter.LabelFrame(self.WIDGETS['LDD_F'], padx=5, pady=5, text='이전 조회 정보')
                    self.WIDGETS['LDD_F_CUR'] = tkinter.LabelFrame(self.WIDGETS['LDD_F'], padx=5, pady=5, text='현재 조회 정보')
                    self.WIDGETS['LDD_F_COM'].place(relheight=1, relwidth=0.5)
                    self.WIDGETS['LDD_F_CUR'].place(relheight=1, relwidth=0.5, relx=0.5)
                else:
                    self._W, self._H = 500, 800
                    if _data['STATUS_FLAG'] == 1:
                        self.WIDGETS['LDD_F_CUR'] = tkinter.LabelFrame(self.WIDGETS['LDD_F'], padx=5, pady=5, text='현재 조회 정보')
                        self.WIDGETS['LDD_F_CUR'].place(relheight=1, relwidth=1)
                    else:
                        self.WIDGETS['LDD_F_CUR'] = tkinter.LabelFrame(self.WIDGETS['LDD_F'], padx=5, pady=5, text='신규 조회 정보')
                        self.WIDGETS['LDD_F_CUR'].place(relheight=1, relwidth=1)
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 상세 정보')
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['LDD_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                if _data['STATUS_FLAG'] == 2:
                    _row = 0
                    for data_set in _data['COM_DATA']:
                        if type(data_set) == dict:
                            for key, value in data_set.items():
                                _label_title = tkinter.Label(self.WIDGETS['LDD_F_COM'], anchor=tkinter.W, width=12, text=key)
                                _label_value = tkinter.Label(self.WIDGETS['LDD_F_COM'], anchor=tkinter.W, width=30, text=value)
                                _label_title.place(relwidth=0.25, height=24, y=_row)
                                _label_value.place(relwidth=0.75, height=24, relx=0.25, y=_row)
                                _label_title.bind('<Button-1>', lambda event, _k=key, _v=value: _copy_command(_event=event, _key=_k, _value=_v))
                                _label_value.bind('<Button-1>', lambda event, _k=key, _v=value: _copy_command(_event=event, _key=_k, _value=_v))
                                _row += 24
                            _row += 6
                            ttk.Separator(self.WIDGETS['LDD_F_COM'], orient=tkinter.HORIZONTAL).place(relwidth=1, height=1, y=_row)
                            _row += 6
                        else:
                            _row += 6
                            frame_table = tkinter.Frame(self.WIDGETS['LDD_F_COM'])
                            tree_y_scroll = tkinter.Scrollbar(frame_table)
                            tree_y_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
                            tree_x_scroll = tkinter.Scrollbar(frame_table, orient=tkinter.HORIZONTAL)
                            tree_x_scroll.pack(side=tkinter.BOTTOM, fill=tkinter.X)
                            self.WIDGETS['T_COM'] = ttk.Treeview(frame_table, show='headings', xscrollcommand=tree_x_scroll.set, yscrollcommand=tree_y_scroll.set, selectmode='none')
                            self.WIDGETS['T_COM'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                            tree_y_scroll.config(command=self.WIDGETS['T_COM'].yview)
                            tree_x_scroll.config(command=self.WIDGETS['T_COM'].xview)
                            frame_table.place(relwidth=1, height=(self._H - _row - 50), y=_row)
                            _tree_columns = ['B/L 번호', '선적일자', '선기적포장개수', '선기적중량']
                            self.WIDGETS['T_COM']['columns'] = ['B/L 번호', '선적일자', '선기적포장개수', '선기적중량']
                            for key in _tree_columns:
                                _col_width = 0
                                if key == 'B/L 번호':
                                    _col_width = 197
                                elif key == '선적일자':
                                    _col_width = 100
                                else:
                                    _col_width = 80
                                self.WIDGETS['T_COM'].column(key, anchor=tkinter.CENTER, width=_col_width, stretch=tkinter.NO)
                                self.WIDGETS['T_COM'].heading(key, text=key)
                            for ti in range(0, len(data_set)):
                                self.WIDGETS['T_COM'].insert(parent='', values=data_set[ti], index=ti)
                            self.WIDGETS['T_COM'].bind('<Button-1>', lambda event: self._event_table_cell_click(_copy_command=_copy_command, _event=event, _code='T_COM'))
                _row = 0
                for i in range(0, len(_data['CUR_DATA'])):
                    if type(_data['CUR_DATA'][i]) == dict:
                        for key, value in _data['CUR_DATA'][i].items():
                            if len(_data['COM_DATA']) != 0 and i != 0:
                                if _data['COM_DATA'][i].get(key) != value:
                                    _label_title = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=12, bg='#9DFFCE', text=key)
                                    _label_title.place(relwidth=0.25, height=24, y=_row)
                                    _label_value = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=30, bg='#9DFFCE', text=value)
                                    _label_value.place(relwidth=0.75, height=24, relx=0.25, y=_row)
                                else:
                                    _label_title = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=12, text=key)
                                    _label_title.place(relwidth=0.25, height=24, y=_row)
                                    _label_value = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=30, text=value)
                                    _label_value.place(relwidth=0.75, height=24, relx=0.25, y=_row)
                            else:
                                _label_title = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=12, text=key)
                                _label_title.place(relwidth=0.25, height=24, y=_row)
                                _label_value = tkinter.Label(self.WIDGETS['LDD_F_CUR'], anchor=tkinter.W, width=30, text=value)
                                _label_value.place(relwidth=0.75, height=24, relx=0.25, y=_row)
                            _label_title.bind('<Button-1>', lambda event, _k=key, _v=value: _copy_command(_event=event, _key=_k, _value=_v))
                            _label_value.bind('<Button-1>', lambda event, _k=key, _v=value: _copy_command(_event=event, _key=_k, _value=_v))
                            _row += 24
                        _row += 6
                        ttk.Separator(self.WIDGETS['LDD_F_CUR'], orient=tkinter.HORIZONTAL).place(relwidth=1, height=1, y=_row)
                        _row += 6
                    else:
                        _row += 6
                        frame_table = tkinter.Frame(self.WIDGETS['LDD_F_CUR'])
                        tree_y_scroll = tkinter.Scrollbar(frame_table)
                        tree_y_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)
                        tree_x_scroll = tkinter.Scrollbar(frame_table, orient=tkinter.HORIZONTAL)
                        tree_x_scroll.pack(side=tkinter.BOTTOM, fill=tkinter.X)
                        self.WIDGETS['T_CUR'] = ttk.Treeview(frame_table, name='tree_search', show='headings', xscrollcommand=tree_x_scroll.set, yscrollcommand=tree_y_scroll.set, selectmode='none')
                        self.WIDGETS['T_CUR'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                        tree_y_scroll.config(command=self.WIDGETS['T_CUR'].yview)
                        tree_x_scroll.config(command=self.WIDGETS['T_CUR'].xview)
                        frame_table.place(relwidth=1, height=(self._H - _row - 50), y=_row)
                        _tree_columns = ['B/L 번호', '선적일자', '선기적포장개수', '선기적중량']
                        self.WIDGETS['T_CUR']['columns'] = ['B/L 번호', '선적일자', '선기적포장개수', '선기적중량']
                        for key in _tree_columns:
                            _col_width = 0
                            if key == 'B/L 번호':
                                _col_width = 197
                            elif key == '선적일자':
                                _col_width = 100
                            else:
                                _col_width = 80
                            self.WIDGETS['T_CUR'].column(key, anchor=tkinter.CENTER, width=_col_width, stretch=tkinter.NO)
                            self.WIDGETS['T_CUR'].heading(key, text=key)
                        for ti in range(0, len(_data['CUR_DATA'][i])):
                            if len(_data['COM_DATA']) != 0:
                                if _data['CUR_DATA'][i][ti] in _data['COM_DATA'][i]:
                                    self.WIDGETS['T_CUR'].insert(parent='', values=_data['CUR_DATA'][i][ti], index=ti)
                                else:
                                    self.WIDGETS['T_CUR'].insert(parent='', values=_data['CUR_DATA'][i][ti], index=ti, tags='changed')
                            else:
                                self.WIDGETS['T_CUR'].insert(parent='', values=_data['CUR_DATA'][i][ti], index=ti)
                        self.WIDGETS['T_CUR'].tag_configure('changed', background='#9DFFCE')
                        self.WIDGETS['T_CUR'].bind('<Button-1>', lambda event: self._event_table_cell_click(_copy_command=_copy_command, _event=event, _code='T_CUR'))
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
            else:
                self.destroy()
                self.detail_data(_data, _copy_command)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_DetailData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '차량 상세 정보 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_excel_check(self, _flag: int = 0):
        try:
            self.APP.APP_QUEUE.put(_flag)
            self.destroy()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickExcelCheck',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '다운로드 엑셀 처리 중 오류가 발생하였습니다.')

    def excel_check(self, _path: str = ''):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Configure>', self._after_covering)
                self.bind('<Escape>', self.destroy)
                self.wm_attributes('-alpha', 0.9)
                self.overrideredirect(True)
                self.resizable(False, False)
                self.title('UNIPASS :: ')
                self.WIDGETS['ODE_F'] = tkinter.Frame(self, padx=5, pady=5)
                self.WIDGETS['ODE_PATH'] = tkinter.Label(self, text=_path, padx=5, pady=5)
                self.WIDGETS['ODE_EXCEL'] = tkinter.Button(self.WIDGETS['ODE_F'], text='엑셀 열기', overrelief=tkinter.SOLID, command=lambda _flag=1: self._event_button_click_excel_check(_flag), repeatdelay=1000, repeatinterval=100, width=20, height=3)
                self.WIDGETS['ODE_DIR'] = tkinter.Button(self.WIDGETS['ODE_F'], text='위치 열기', overrelief=tkinter.SOLID, command=lambda _flag=2: self._event_button_click_excel_check(_flag), repeatdelay=1000, repeatinterval=100, width=20, height=3)
                self.WIDGETS['ODE_CANCEL'] = tkinter.Button(self.WIDGETS['ODE_F'], text='취소', overrelief=tkinter.SOLID, command=self._event_button_click_excel_check, repeatdelay=1000, repeatinterval=100, width=20, height=3)
                self.WIDGETS['ODE_PATH'].pack(anchor=tkinter.S, expand=tkinter.TRUE, pady=5)
                self.WIDGETS['ODE_EXCEL'].pack(side=tkinter.LEFT)
                self.WIDGETS['ODE_DIR'].pack(side=tkinter.LEFT, padx=20)
                self.WIDGETS['ODE_CANCEL'].pack(side=tkinter.LEFT)
                self.WIDGETS['ODE_F'].pack(anchor=tkinter.N, expand=tkinter.TRUE, pady=5)
                self.lift()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ExcelCheck',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '다운로드 엑셀 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_add_manager(self):
        try:
            _add_id = self.WIDGETS['AM_TE_ID'].get_value()
            _add_pw = self.WIDGETS['AM_TE_PW'].get_value()
            _add_pw_check = self.WIDGETS['AM_TE_PW_CHECK'].get_value()
            _add_name = self.WIDGETS['AM_TE_NAME'].get_value()
            if _add_id != '' and _add_pw != '' and _add_pw_check != '' and _add_name != '':
                if _add_pw == _add_pw_check:
                    self.APP.APP_QUEUE.put({
                        'ID': _add_id,
                        'PW': _add_pw,
                        'NAME': _add_name
                    })
                    self.destroy()
                else:
                    messagebox.showwarning('UNIPASS ::', '입력된 두개의 비밀번호가 다릅니다.\n다시 입력해주세요.')
            else:
                messagebox.showwarning('UNIPASS ::', '입력되지 않은 항목이 있습니다.\n모두 입력해주세요.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickAddManager',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 추가 중 오류가 발생하였습니다.')

    def add_manager(self, _prev_manager_info: [dict, None]):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 400, 292
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 계정 추가')
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['AM_F'] = tkinter.Frame(self, padx=10, pady=10)
                self.WIDGETS['AM_TE_ID'] = TitleEntry(self.WIDGETS['AM_F'], '아이디', _label_size=3)
                self.WIDGETS['AM_TE_PW'] = TitleEntry(self.WIDGETS['AM_F'], '비밀번호', True, _label_size=3)
                self.WIDGETS['AM_TE_PW_CHECK'] = TitleEntry(self.WIDGETS['AM_F'], '비밀번호 재확인', True, _label_size=3)
                self.WIDGETS['AM_TE_NAME'] = TitleEntry(self.WIDGETS['AM_F'], '이름', _label_size=3)
                self.WIDGETS['AM_B_ADD'] = tkinter.Button(self.WIDGETS['AM_F'], text='추가', overrelief=tkinter.SOLID, command=self._event_button_click_add_manager, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['AM_TE_ID'].place(relwidth=1, height=44)
                self.WIDGETS['AM_TE_PW'].place(y=54, relwidth=1, height=44)
                self.WIDGETS['AM_TE_PW_CHECK'].place(y=108, relwidth=1, height=44)
                self.WIDGETS['AM_TE_NAME'].place(y=162, relwidth=1, height=44)
                self.WIDGETS['AM_B_ADD'].place(y=216, relwidth=1, height=56)
                self.WIDGETS['AM_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                if _prev_manager_info is not None:
                    self.WIDGETS['AM_TE_ID'].set_value(_prev_manager_info['ID'])
                    self.WIDGETS['AM_TE_PW'].set_value(_prev_manager_info['PW'])
                    self.WIDGETS['AM_TE_PW_CHECK'].set_value(_prev_manager_info['PW'])
                    self.WIDGETS['AM_TE_NAME'].set_value(_prev_manager_info['NAME'])
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_AddManager',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 추가 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_manager_menu(self, _menu_index: int):
        try:
            self.APP.APP_QUEUE.put(_menu_index)
            self.destroy()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickManagerMenu',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 관리 메뉴 처리 중 오류가 발생하였습니다.')

    def manager_menu(self, _manager_id: str):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 300, 152
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 계정 관리'.format(_manager_id))
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['MM_L_ID'] = tkinter.Label(self, text=_manager_id, anchor=tkinter.CENTER)
                self.WIDGETS['MM_B_PW_CHANGE'] = tkinter.Button(self, text='비밀번호 변경', overrelief=tkinter.SOLID, command=lambda _mi=1: self._event_button_click_manager_menu(_menu_index=_mi), repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['MM_B_ID_DELETE'] = tkinter.Button(self, text='계정 삭제', overrelief=tkinter.SOLID, command=lambda _mi=2: self._event_button_click_manager_menu(_menu_index=_mi), repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['MM_L_ID'].place(x=10, width=self._W - 20, height=24, y=10)
                self.WIDGETS['MM_B_PW_CHANGE'].place(x=10, width=self._W - 20, height=44, y=44)
                self.WIDGETS['MM_B_ID_DELETE'].place(x=10, width=self._W - 20, height=44, y=98)
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
            else:
                self.destroy()
                self.manager_menu(_manager_id)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ManagerMenu',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 관리 메뉴 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_change_manager_password(self):
        try:
            _change_pw = self.WIDGETS['CMP_TE_PW'].get_value()
            _change_pw_check = self.WIDGETS['CMP_TE_PW_CHECK'].get_value()
            if _change_pw != '' and _change_pw_check != '':
                if _change_pw == _change_pw_check:
                    self.APP.APP_QUEUE.put(_change_pw)
                    self.destroy()
                else:
                    messagebox.showwarning('UNIPASS ::', '입력된 두개의 비밀번호가 다릅니다.\n다시 입력해주세요.')
            else:
                messagebox.showwarning('UNIPASS ::', '입력되지 않은 항목이 있습니다.\n모두 입력해주세요.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickChangeManagerPassword',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 추가 중 오류가 발생하였습니다.')

    def change_manager_password(self):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 400, 184
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 계정 비밀번호 변경')
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['CMP_F'] = tkinter.Frame(self, padx=10, pady=10)
                self.WIDGETS['CMP_TE_PW'] = TitleEntry(self.WIDGETS['CMP_F'], '비밀번호', True, _label_size=3)
                self.WIDGETS['CMP_TE_PW_CHECK'] = TitleEntry(self.WIDGETS['CMP_F'], '비밀번호 재확인', True, _label_size=3)
                self.WIDGETS['CMP_B_CHANGE'] = tkinter.Button(self.WIDGETS['CMP_F'], text='변경', overrelief=tkinter.SOLID, command=self._event_button_click_change_manager_password, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['CMP_TE_PW'].place(relwidth=1, height=44)
                self.WIDGETS['CMP_TE_PW_CHECK'].place(y=54, relwidth=1, height=44)
                self.WIDGETS['CMP_B_CHANGE'].place(y=108, relwidth=1, height=56)
                self.WIDGETS['CMP_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ChangeManagerPassword',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '계정 비밀번호 변경 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_bl_detail_print(self):
        self.APP.APP_QUEUE.put(self.WIDGETS['LC_PRINT'].get_value_text())
        self.destroy()

    def _event_button_click_prev_bl_detail(self):
        self.APP.APP_QUEUE.put(False)
        self.destroy()

    def _event_button_click_next_bl_detail(self):
        self.APP.APP_QUEUE.put(True)
        self.destroy()

    def bl_list_print(self, _head_data, _body_data, _copy_command, _move_button_status):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 1040, 700
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: B/L 번호 조회 결과 프린트')
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['BLP_F'] = tkinter.Frame(self, padx=10, pady=10)
                _print_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                category_printer = {win32print.GetDefaultPrinter(): 0}
                for _printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, None, 2):
                    _category_no = 1
                    if 'microsoft' not in str(_printer['pDriverName']).lower() and str(_printer['pPrinterName']) not in category_printer:
                        category_printer.__setitem__(_printer['pPrinterName'], _category_no)
                        _category_no += 1
                self.WIDGETS['B_PREV'] = tkinter.Button(self.WIDGETS['BLP_F'], text='이전', overrelief=tkinter.SOLID, command=self._event_button_click_prev_bl_detail, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['B_NEXT'] = tkinter.Button(self.WIDGETS['BLP_F'], text='다음', overrelief=tkinter.SOLID, command=self._event_button_click_next_bl_detail, repeatdelay=1000, repeatinterval=100)
                if not _move_button_status[0]:
                    self.WIDGETS['B_PREV'].configure(state=tkinter.DISABLED)
                if not _move_button_status[1]:
                    self.WIDGETS['B_NEXT'].configure(state=tkinter.DISABLED)
                _total_amount = 0
                _total_weight = 0
                self.WIDGETS['LC_PRINT'] = LabelCombo(self.WIDGETS['BLP_F'], '프린터 선택', category_printer)
                self.WIDGETS['B_PRINT'] = tkinter.Button(self.WIDGETS['BLP_F'], text='프린트', overrelief=tkinter.SOLID, command=self._event_button_click_bl_detail_print, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['L_PRINT_DATE'] = tkinter.Label(self.WIDGETS['BLP_F'], text='출력일자', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_PRINT_DATE'] = tkinter.Label(self.WIDGETS['BLP_F'], text=_print_date, borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['L_LOOKUP_DATE'] = tkinter.Label(self.WIDGETS['BLP_F'], text='조회일자', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_LOOKUP_DATE'] = tkinter.Label(self.WIDGETS['BLP_F'], text=_head_data[0], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['L_BL_NUMBER'] = tkinter.Label(self.WIDGETS['BLP_F'], text='B/L 번호', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_BL_NUMBER'] = tkinter.Label(self.WIDGETS['BLP_F'], text=_head_data[1], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['L_TOTAL_COUNT'] = tkinter.Label(self.WIDGETS['BLP_F'], text='면장 개수', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_TOTAL_COUNT'] = tkinter.Label(self.WIDGETS['BLP_F'], text=str(len(_body_data)) if _body_data[0][0] is not None else '0', borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['L_TOTAL_AMOUNT'] = tkinter.Label(self.WIDGETS['BLP_F'], text='선적 대수', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_TOTAL_AMOUNT'] = tkinter.Label(self.WIDGETS['BLP_F'], text=str(_total_amount), borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['L_TOTAL_WEIGHT'] = tkinter.Label(self.WIDGETS['BLP_F'], text='선적 중량', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray')
                self.WIDGETS['V_TOTAL_WEIGHT'] = tkinter.Label(self.WIDGETS['BLP_F'], text=str(_total_weight), borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                self.WIDGETS['B_PREV'].place(height=56, relwidth=0.15)
                self.WIDGETS['LC_PRINT'].place(height=56, relwidth=0.322, relx=0.26)
                self.WIDGETS['B_PRINT'].place(height=56, relwidth=0.15, relx=0.591)
                self.WIDGETS['B_NEXT'].place(height=56, relwidth=0.15, relx=0.851)
                self.WIDGETS['L_PRINT_DATE'].place(height=33, relwidth=0.22, y=66)
                self.WIDGETS['V_PRINT_DATE'].place(height=33, relwidth=0.22, y=99)
                self.WIDGETS['L_LOOKUP_DATE'].place(height=33, relwidth=0.22, y=66, relx=0.26)
                self.WIDGETS['V_LOOKUP_DATE'].place(height=33, relwidth=0.22, y=99, relx=0.26)
                self.WIDGETS['L_BL_NUMBER'].place(height=33, relwidth=0.22, y=66, relx=0.52)
                self.WIDGETS['V_BL_NUMBER'].place(height=33, relwidth=0.22, y=99, relx=0.52)
                self.WIDGETS['L_TOTAL_COUNT'].place(height=33, relwidth=0.08, y=66, relx=0.78)
                self.WIDGETS['V_TOTAL_COUNT'].place(height=33, relwidth=0.08, y=99, relx=0.78)
                self.WIDGETS['L_TOTAL_AMOUNT'].place(height=33, relwidth=0.07, y=66, relx=0.86)
                self.WIDGETS['V_TOTAL_AMOUNT'].place(height=33, relwidth=0.07, y=99, relx=0.86)
                self.WIDGETS['L_TOTAL_WEIGHT'].place(height=33, relwidth=0.07, y=66, relx=0.93)
                self.WIDGETS['V_TOTAL_WEIGHT'].place(height=33, relwidth=0.07, y=99, relx=0.93)
                self.WIDGETS['L_PRINT_DATE'].bind('<Button-1>', lambda _event: _copy_command(_event, '출력일자', _print_date))
                self.WIDGETS['V_PRINT_DATE'].bind('<Button-1>', lambda _event: _copy_command(_event, '출력일자', _print_date))
                self.WIDGETS['L_LOOKUP_DATE'].bind('<Button-1>', lambda _event: _copy_command(_event, '조회일자', _head_data[0]))
                self.WIDGETS['V_LOOKUP_DATE'].bind('<Button-1>', lambda _event: _copy_command(_event, '조회일자', _head_data[0]))
                self.WIDGETS['L_BL_NUMBER'].bind('<Button-1>', lambda _event: _copy_command(_event, 'B/L 번호', _head_data[1]))
                self.WIDGETS['V_BL_NUMBER'].bind('<Button-1>', lambda _event: _copy_command(_event, 'B/L 번호', _head_data[1]))
                self.WIDGETS['L_TOTAL_COUNT'].bind('<Button-1>', lambda _event: _copy_command(_event, '면장 개수', str(len(_body_data))))
                self.WIDGETS['V_TOTAL_COUNT'].bind('<Button-1>', lambda _event: _copy_command(_event, '면장 개수', str(len(_body_data))))
                self.WIDGETS['L_TOTAL_AMOUNT'].bind('<Button-1>', lambda _event: _copy_command(_event, '선적 개수', str(_total_amount)))
                self.WIDGETS['V_TOTAL_AMOUNT'].bind('<Button-1>', lambda _event: _copy_command(_event, '선적 개수', str(_total_amount)))
                self.WIDGETS['L_TOTAL_WEIGHT'].bind('<Button-1>', lambda _event: _copy_command(_event, '선적 중량', str(_total_weight)))
                self.WIDGETS['V_TOTAL_WEIGHT'].bind('<Button-1>', lambda _event: _copy_command(_event, '선적 중량', str(_total_weight)))
                self.WIDGETS['BLP_SCR_F'] = tkinter.Frame(self.WIDGETS['BLP_F'])
                if len(_body_data) > 0 and _body_data[0][0] is not None:
                    self.WIDGETS['BLP_CANVAS'] = tkinter.Canvas(self.WIDGETS['BLP_SCR_F'])
                    self.WIDGETS['BLP_SCROLL'] = ttk.Scrollbar(self.WIDGETS['BLP_SCR_F'], orient=tkinter.VERTICAL, command=self.WIDGETS['BLP_CANVAS'].yview)
                    self.WIDGETS['BLP_CONTAINER'] = tkinter.Frame(self.WIDGETS['BLP_CANVAS'], background='white')
                    self.WIDGETS['BLP_CANVAS'].bind_all('<MouseWheel>', lambda _event: self.WIDGETS['BLP_CANVAS'].yview_scroll(int(-1 * (_event.delta / 120)), "units"))
                    self.WIDGETS['BLP_CONTAINER'].bind('<Configure>', lambda _event: self.WIDGETS['BLP_CANVAS'].configure(scrollregion=self.WIDGETS['BLP_CANVAS'].bbox('all')))
                    self.WIDGETS['BLP_CANVAS'].create_window((0, len(_body_data) * 100), window=self.WIDGETS['BLP_CONTAINER'], anchor=tkinter.CENTER)
                    self.WIDGETS['BLP_CANVAS'].configure(yscrollcommand=self.WIDGETS['BLP_SCROLL'].set)
                    for _row_index in range(len(_body_data)):
                        _ = tkinter.Frame(self.WIDGETS['BLP_CONTAINER'], height=165, width=1000)
                        if _row_index % 2 != 0 and _row_index != 0:
                            _.pack(fill=tkinter.X, side=tkinter.TOP, pady=30)
                        else:
                            _.pack(fill=tkinter.X, side=tkinter.TOP)
                        tkinter.Label(_, text='통관사항', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray').place(relwidth=0.44, height=33)
                        tkinter.Label(_, text='선적사항', borderwidth=1, relief=tkinter.SOLID, background='#DDDDDD', foreground='gray').place(relwidth=0.56, height=33, relx=0.44)
                        _l_es = tkinter.Label(_, text='수출자', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_es.place(relwidth=0.2, height=33, y=33)
                        _l_ad = tkinter.Label(_, text='수리일자', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_ad.place(relwidth=0.12, height=33, y=33, relx=0.2)
                        _l_ta = tkinter.Label(_, text='통관포장개수', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_ta.place(relwidth=0.12, height=33, y=33, relx=0.32)
                        _l_sp = tkinter.Label(_, text='선기적지', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_sp.place(relwidth=0.12, height=33, y=33, relx=0.64)
                        _l_pa = tkinter.Label(_, text='선기적포장개수', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_pa.place(relwidth=0.12, height=33, y=33, relx=0.76)
                        _l_cc = tkinter.Label(_, text='분할회수', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_cc.place(relwidth=0.12, height=33, y=33, relx=0.88)
                        _l_en = tkinter.Label(_, text='수출신고번호', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_en.place(relwidth=0.2, height=33, y=66)
                        _l_ldd = tkinter.Label(_, text='적재의무기한', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_ldd.place(relwidth=0.12, height=33, y=66, relx=0.2)
                        _l_tw = tkinter.Label(_, text='통관중량', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_tw.place(relwidth=0.12, height=33, y=66, relx=0.32)
                        _l_dd = tkinter.Label(_, text='출항일자', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_dd.place(relwidth=0.12, height=33, y=66, relx=0.64)
                        _l_pw = tkinter.Label(_, text='선기적중량', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_pw.place(relwidth=0.12, height=33, y=66, relx=0.76)
                        _l_fps = tkinter.Label(_, text='선기적완료여부', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_fps.place(relwidth=0.12, height=33, y=66, relx=0.88)
                        _l_mn = tkinter.Label(_, text='적하목록관리번호', borderwidth=1, relief=tkinter.SOLID, background='#EEEEEE', foreground='gray')
                        _l_mn.place(relwidth=0.2, height=66, y=33, relx=0.44)
                        _v_es = tkinter.Label(_, text=_body_data[_row_index][0], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_es.place(relwidth=0.2, height=33, y=99)
                        _v_ad = tkinter.Label(_, text=_body_data[_row_index][1], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_ad.place(relwidth=0.12, height=33, y=99, relx=0.2)
                        _v_ta = tkinter.Label(_, text=f'{_body_data[_row_index][2]:,}', borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_ta.place(relwidth=0.12, height=33, y=99, relx=0.32)
                        _v_sp = tkinter.Label(_, text=_body_data[_row_index][3], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_sp.place(relwidth=0.12, height=33, y=99, relx=0.64)
                        _v_pa = tkinter.Label(_, text=f'{_body_data[_row_index][4]:,}', borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _total_amount += _body_data[_row_index][4]
                        _v_pa.place(relwidth=0.12, height=33, y=99, relx=0.76)
                        _v_cc = tkinter.Label(_, text=_body_data[_row_index][5], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_cc.place(relwidth=0.12, height=33, y=99, relx=0.88)
                        _v_en = tkinter.Label(_, text=_body_data[_row_index][6], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_en.place(relwidth=0.2, height=33, y=132)
                        _v_ldd = tkinter.Label(_, text=_body_data[_row_index][7], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_ldd.place(relwidth=0.12, height=33, y=132, relx=0.2)
                        _v_tw = tkinter.Label(_, text=f'{_body_data[_row_index][8]:,}', borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_tw.place(relwidth=0.12, height=33, y=132, relx=0.32)
                        _v_dd = tkinter.Label(_, text=_body_data[_row_index][9], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_dd.place(relwidth=0.12, height=33, y=132, relx=0.64)
                        _v_pw = tkinter.Label(_, text=f'{_body_data[_row_index][10]:,}', borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _total_weight += _body_data[_row_index][10]
                        _v_pw.place(relwidth=0.12, height=33, y=132, relx=0.76)
                        _v_fps = tkinter.Label(_, text=_body_data[_row_index][11], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_fps.place(relwidth=0.12, height=33, y=132, relx=0.88)
                        _v_mn = tkinter.Label(_, text=_body_data[_row_index][12], borderwidth=1, relief=tkinter.SOLID, background='white', foreground='black')
                        _v_mn.place(relwidth=0.2, height=66, y=99, relx=0.44)
                        _l_es.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][0]: _copy_command(_event, '수출자', _value))
                        _v_es.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][0]: _copy_command(_event, '수출자', _value))
                        _l_ad.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][1]: _copy_command(_event, '수리일자', _value))
                        _v_ad.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][1]: _copy_command(_event, '수리일자', _value))
                        _l_ta.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][2]: _copy_command(_event, '통관포장개수', _value))
                        _v_ta.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][2]: _copy_command(_event, '통관포장개수', _value))
                        _l_sp.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][3]: _copy_command(_event, '선기적지', _value))
                        _v_sp.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][3]: _copy_command(_event, '선기적지', _value))
                        _l_pa.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][4]: _copy_command(_event, '선기적포장개수', _value))
                        _v_pa.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][4]: _copy_command(_event, '선기적포장개수', _value))
                        _l_cc.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][5]: _copy_command(_event, '분할회수', _value))
                        _v_cc.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][5]: _copy_command(_event, '분할회수', _value))
                        _l_en.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][6]: _copy_command(_event, '수출신고번호', _value))
                        _v_en.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][6]: _copy_command(_event, '수출신고번호', _value))
                        _l_ldd.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][7]: _copy_command(_event, '적재의무기한', _value))
                        _v_ldd.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][7]: _copy_command(_event, '적재의무기한', _value))
                        _l_tw.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][8]: _copy_command(_event, '통관중량', _value))
                        _v_tw.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][8]: _copy_command(_event, '통관중량', _value))
                        _l_dd.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][9]: _copy_command(_event, '출항일자', _value))
                        _v_dd.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][9]: _copy_command(_event, '출항일자', _value))
                        _l_pw.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][10]: _copy_command(_event, '선기적중량', _value))
                        _v_pw.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][10]: _copy_command(_event, '선기적중량', _value))
                        _l_fps.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][11]: _copy_command(_event, '선기적완료여부', _value))
                        _v_fps.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][11]: _copy_command(_event, '선기적완료여부', _value))
                        _l_mn.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][12]: _copy_command(_event, '적하목록관리번호', _value))
                        _v_mn.bind('<Button-1>', lambda _event, _value=_body_data[_row_index][12]: _copy_command(_event, '적하목록관리번호', _value))
                    self.WIDGETS['BLP_CANVAS'].pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=tkinter.TRUE)
                    self.WIDGETS['BLP_SCROLL'].pack(fill=tkinter.Y, side=tkinter.RIGHT)
                else:
                    tkinter.Label(self.WIDGETS['BLP_SCR_F'], text='조회 결과가 없습니다.', foreground='gray', background='white').pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                    self.WIDGETS['B_PRINT'].configure(state=tkinter.DISABLED)
                self.WIDGETS['V_TOTAL_AMOUNT'].configure(text=f'{_total_amount:,}')
                self.WIDGETS['V_TOTAL_WEIGHT'].configure(text=f'{_total_weight:,}')
                self.WIDGETS['BLP_SCR_F'].place(x=0, y=142, relwidth=1, height=self._H - 162)
                self.WIDGETS['BLP_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_BlListPrint',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', 'B/L 번호 프린트 팝업 생성 중 오류가 발생하였습니다.')

    def _event_button_click_add_excel_title(self):
        self.APP.APP_QUEUE.put(self.WIDGETS['AM_TE_ID'].get_value())
        self.destroy()

    def add_excel_title(self):
        try:
            if not self.winfo_exists():
                self.initial()
                self.bind('<Escape>', self.destroy)
                self._W, self._H = 400, 130
                _w, _h, _x, _y = self.APP.get_current_geometry()
                if self.APP.state() == 'normal':
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.WW / 2) - int(self._W / 2) + _x, int(self.APP.WH / 2) - int(self._H / 2) + _y))
                else:
                    self.geometry('{}x{}+{}+{}'.format(self._W, self._H, int(self.APP.winfo_width() / 2) - int(self._W / 2), int(self.APP.winfo_height() / 2) - int(self._H / 2)))
                self.wm_attributes('-alpha', 1)
                self.overrideredirect(False)
                self.resizable(False, False)
                self.title('UNIPASS :: 열 제목 추가')
                self.iconbitmap(self.APP.PATH['IC'])
                self.WIDGETS['AM_F'] = tkinter.Frame(self, padx=10, pady=10)
                self.WIDGETS['AM_TE_ID'] = TitleEntry(self.WIDGETS['AM_F'], '열 제목', _label_size=3)
                self.WIDGETS['AM_B_ADD'] = tkinter.Button(self.WIDGETS['AM_F'], text='추가', overrelief=tkinter.SOLID, command=self._event_button_click_add_excel_title, repeatdelay=1000, repeatinterval=100)
                self.WIDGETS['AM_TE_ID'].place(relwidth=1, height=44)
                self.WIDGETS['AM_B_ADD'].place(y=54, relwidth=1, height=56)
                self.WIDGETS['AM_F'].pack(fill=tkinter.BOTH, expand=tkinter.TRUE)
                self.lift()
                self.focus_set()
                self.APP.wait_window(self)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_AddExcelTitle',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS ::', '엑셀 열 제목 추가 팝업 생성 중 오류가 발생하였습니다.')
