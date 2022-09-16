import os
import sys
import queue
import tkinter
import datetime
from tkinter import messagebox
from View import Body, LabelText, LabelCombo, LabelCalendar, Table
from Model.model_common import DownloadBlDataToExcel, PrintBlDetail, DownloadBlDataToPrintForm
from Model.model_lookup_bl_history import GetCategory, GetLookUpList, GetBlList, GetBlDetailList


class LookUpBlHistoryBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewLookUpBlHistory_LookUpBlHistoryBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA['CATEGORY_MAN_ID'] = {'All': 0}
            self.DATA['BL_DETAIL_DATA'] = None
            self.DATA['KEY_INDEX_LIST'] = []
            self.WIDGETS.clear()
            self.request_get_category()
            self.WIDGETS['LOOKUP_LC_ID'] = LabelCombo(self, '조회 아이디', self.DATA['CATEGORY_MAN_ID'])
            self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'] = LabelCalendar(self, '조회일자')
            self.WIDGETS['LOOKUP_B_RESET'] = tkinter.Button(self, text='검색 필터\n초기화', overrelief=tkinter.SOLID, command=self.event_button_click_reset, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['LOOKUP_B_SEARCH'] = tkinter.Button(self, text='검색', overrelief=tkinter.SOLID, command=self.request_get_lookup_list, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['LOOKUP_TABLE'] = Table(self.APP, self)
            self.WIDGETS['BL_TABLE'] = Table(self.APP, self)
            self.WIDGETS['BL_B_PREV'] = tkinter.Button(self, text='뒤로', overrelief=tkinter.SOLID, command=self.event_button_click_prev, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['BL_LT_TOTAL'] = LabelText(self, '전체 개수', self.WIDGETS['BL_TABLE'].COUNT['TOT'])
            self.WIDGETS['BL_B_COPY'] = tkinter.Button(self, textvariable=self.WIDGETS['BL_TABLE'].CELL_SELECTION_COUNT, overrelief=tkinter.SOLID, command=self.event_button_click_selection_copy, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['BL_B_DOWNLOAD'] = tkinter.Button(self, text='목록 엑셀\n다운로드', overrelief=tkinter.SOLID, command=self.event_button_click_excel_download, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['BL_B_PF_DOWNLOAD'] = tkinter.Button(self, text='양식 엑셀\n다운로드', overrelief=tkinter.SOLID, command=self.event_button_click_print_form_download, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['LOOKUP_TABLE'].set_table_resizing()
            self.WIDGETS['LOOKUP_TABLE'].set_table_paging()
            self.WIDGETS['LOOKUP_TABLE'].set_table_celling()
            self.WIDGETS['LOOKUP_TABLE'].set_head_sort_command()
            self.WIDGETS['LOOKUP_TABLE'].set_cell_double_click_command()
            self.WIDGETS['LOOKUP_TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.COMMON_HEAD_CATEGORY.items() if 'LHL' in _value['TABLE']},
                _is_change=True
            )
            self.WIDGETS['LOOKUP_TABLE'].cell_clear()
            self.WIDGETS['LOOKUP_TABLE'].clear_count()
            self.WIDGETS['LOOKUP_TABLE'].delete_row()
            self.WIDGETS['LOOKUP_TABLE'].bind('<Double-1>', self.event_table_double_click)
            self.WIDGETS['LOOKUP_TABLE'].set_location_y(86)
            self.WIDGETS['BL_TABLE'].set_table_resizing(False)
            self.WIDGETS['BL_TABLE'].set_table_paging()
            self.WIDGETS['BL_TABLE'].set_table_celling(True)
            self.WIDGETS['BL_TABLE'].set_head_sort_command(_sort_command=self.event_table_head_click)
            self.WIDGETS['BL_TABLE'].set_cell_double_click_command(_double_click_command=self.event_table_cell_double_click)
            self.WIDGETS['BL_TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.BL_HEAD_CATEGORY.items() if 'LHBL' in _value['TABLE']},
                _is_change=True
            )
            self.WIDGETS['BL_TABLE'].head_sort_clear()
            self.WIDGETS['BL_TABLE'].cell_clear()
            self.WIDGETS['BL_TABLE'].clear_count()
            self.WIDGETS['BL_TABLE'].delete_row()
            self.WIDGETS['BL_TABLE'].set_location_y(152)
            self.view_lookup_list()
            self.request_get_lookup_list()
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

    def view_lookup_list(self):
        try:
            self.WIDGETS['BL_B_PREV'].place_forget()
            self.WIDGETS['BL_LT_TOTAL'].place_forget()
            self.WIDGETS['BL_B_COPY'].place_forget()
            self.WIDGETS['BL_B_DOWNLOAD'].place_forget()
            self.WIDGETS['BL_B_PF_DOWNLOAD'].place_forget()
            self.APP.unbind('<Control-c>')
            self.WIDGETS['BL_TABLE'].set_table_resizing(False)
            self.WIDGETS['BL_TABLE'].head_sort_clear()
            self.WIDGETS['BL_TABLE'].cell_clear()
            self.WIDGETS['BL_TABLE'].clear_count()
            self.WIDGETS['BL_TABLE'].delete_row()
            self.WIDGETS['BL_TABLE'].LOCATION.place_forget()
            self.WIDGETS['LOOKUP_LC_ID'].place(height=56, relwidth=0.24325)
            self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'].place(height=56, relwidth=0.24325, relx=0.25225)
            self.WIDGETS['LOOKUP_B_RESET'].place(height=56, relwidth=0.06, relx=0.871)
            self.WIDGETS['LOOKUP_B_SEARCH'].place(height=56, relwidth=0.06, relx=0.941)
            self.WIDGETS['LOOKUP_TABLE'].set_table_resizing()
            self.WIDGETS['LOOKUP_TABLE'].LOCATION.place(y=66, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 86)
            self.DATA['CURRENT_I_INDEX'] = 0
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ViewLookUpList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '화면 초기화 중 오류가 발생하였습니다.')

    def view_bl_list(self):
        try:
            self.WIDGETS['LOOKUP_LC_ID'].place_forget()
            self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'].place_forget()
            self.WIDGETS['LOOKUP_B_RESET'].place_forget()
            self.WIDGETS['LOOKUP_B_SEARCH'].place_forget()
            self.WIDGETS['LOOKUP_TABLE'].set_table_resizing(False)
            self.WIDGETS['LOOKUP_TABLE'].LOCATION.place_forget()
            self.WIDGETS['BL_B_PREV'].place(height=56, relwidth=0.06)
            self.WIDGETS['BL_LT_TOTAL'].place(height=56, relwidth=0.07, y=66)
            self.WIDGETS['BL_B_COPY'].place(height=56, relwidth=0.06, relx=0.803, y=66)
            self.WIDGETS['BL_B_DOWNLOAD'].place(height=56, relwidth=0.06, relx=0.872, y=66)
            self.WIDGETS['BL_B_PF_DOWNLOAD'].place(height=56, relwidth=0.06, relx=0.941, y=66)
            self.APP.bind('<Control-c>', self.event_button_click_selection_copy)
            self.WIDGETS['BL_TABLE'].set_table_resizing(True)
            self.WIDGETS['BL_TABLE'].LOCATION.place(y=132, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 152)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ViewBlList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '화면 초기화 중 오류가 발생하였습니다.')

    def event_button_click_reset(self):
        self.WIDGETS['LOOKUP_LC_ID'].clear_value()
        self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'].clear_value()

    def event_table_double_click(self, _event=None):
        try:
            self.DATA['SELECTED_LOOKUP_INDEX'] = list(self.WIDGETS['LOOKUP_TABLE'].item(self.WIDGETS['LOOKUP_TABLE'].selection()[0]).values())[2][0]
            self.request_get_bl_list()
            self.view_bl_list()
        except IndexError:
            pass
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
            messagebox.showerror('UNIPASS :: ', '차량 목록을 불러오는 중 오류가 발생하였습니다.')

    def event_button_click_prev(self):
        self.view_lookup_list()

    def event_button_click_selection_copy(self, _event=None):
        _tsc = None
        try:
            if len(self.WIDGETS['BL_TABLE'].CELL_SELECTIONS) != 0:
                _tsc = self.WIDGETS['BL_TABLE'].selection_copy()
                if _tsc[0]:
                    self.APP.clipboard_clear()
                    self.APP.clipboard_append(_tsc[1])
                    self.APP.SHORTS.show('선택된 셀이 복사되었습니다.')
                else:
                    _tsc['ECD'] = self._code + _tsc['ECD']
                    self.APP.request_error_report_db(_error_info=_tsc[1])
                    messagebox.showerror('UNIPASS :: ', '복사에 실패하였습니다.')
            else:
                self.APP.SHORTS.show('선택된 셀이 없습니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickSelectionCopy',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_event, _tsc]
                }
            )
            messagebox.showerror('UNIPASS :: ', '복사에 실패하였습니다.')

    def event_button_click_excel_download(self):
        _download_excel = ''
        try:
            self.request_download_data_to_excel()
            _download_excel = self.APP.APP_QUEUE.get()
            self.APP.POPUP.excel_check(_download_excel)
            _open_flag = self.APP.APP_QUEUE.get()
            if _open_flag == 1:
                os.startfile(os.path.realpath(_download_excel))
            elif _open_flag == 2:
                os.startfile(os.path.realpath('/'.join(_download_excel.split('/')[:-1])))
        except queue.Empty:
            pass
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickExcelDownload',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_download_excel]
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')

    def event_button_click_print_form_download(self):
        _download_excel = ''
        try:
            self.request_download_data_to_print_form()
            _download_excel = self.APP.APP_QUEUE.get()
            self.APP.POPUP.excel_check(_download_excel)
            _open_flag = self.APP.APP_QUEUE.get()
            if _open_flag == 1:
                os.startfile(os.path.realpath(_download_excel))
            elif _open_flag == 2:
                os.startfile(os.path.realpath('/'.join(_download_excel.split('/')[:-1])))
        except queue.Empty:
            pass
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickExcelDownload',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_download_excel]
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')

    def event_table_head_click(self, _head_key: str = ''):
        try:
            self.WIDGETS['BL_TABLE'].head_sort(_head_key)
            self.WIDGETS['BL_TABLE'].cell_clear()
            self.WIDGETS['BL_TABLE'].clear_count()
            self.WIDGETS['BL_TABLE'].delete_row()
            self.request_get_bl_list()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventTableHeadClick',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_head_key]
                }
            )
            messagebox.showerror('UNIPASS :: ', '테이블 정렬 중 오류가 발생하였습니다.')

    def event_table_cell_double_click(self, _ici: int):
        try:
            self.DATA['BL_DETAIL_DATA'] = None
            self.APP.APP_QUEUE.put(_ici)
            self.request_get_bl_detail_list()
            _move_button_status = [True, True]
            if self.DATA['KEY_INDEX_LIST'].index(_ici) == 0:
                _move_button_status[0] = False
            if self.DATA['KEY_INDEX_LIST'].index(_ici) + 1 == len(self.DATA['KEY_INDEX_LIST']):
                _move_button_status[1] = False
            if self.DATA['BL_DETAIL_DATA'] is not None:
                self.APP.POPUP.bl_list_print(
                    _head_data=[self.DATA['BL_DETAIL_DATA'][0][0], self.DATA['BL_DETAIL_DATA'][0][1]],
                    _body_data=[_row[2:] for _row in self.DATA['BL_DETAIL_DATA']],
                    _copy_command=self.event_popup_cell_click,
                    _move_button_status=_move_button_status
                )
                if not self.APP.APP_QUEUE.empty():
                    _flag = self.APP.APP_QUEUE.get()
                    if type(_flag) is not bool:
                        self.APP.APP_QUEUE.put(_flag)
                        self.request_bl_detail_print()
                    else:
                        if _flag:
                            self.event_table_cell_double_click(
                                self.DATA['KEY_INDEX_LIST'][
                                    self.DATA['KEY_INDEX_LIST'].index(_ici) + 1
                                    ]
                            )
                        else:
                            self.event_table_cell_double_click(
                                self.DATA['KEY_INDEX_LIST'][
                                    self.DATA['KEY_INDEX_LIST'].index(_ici) - 1
                                    ]
                            )
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventTableCellDoubleClick',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '상세 정보를 불러오는 중 오류가 발생하였습니다.')

    def event_popup_cell_click(self, _event, _key, _value):
        try:
            self.APP.clipboard_clear()
            self.APP.clipboard_append(_value)
            self.APP.SHORTS.show('{} 이(가) 복사되었습니다.'.format(_key))
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventPopupCellClick',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': [_key, _value]
                }
            )
            messagebox.showerror('UNIPASS :: ', '복사에 실패하였습니다.')

    def request_get_category(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_category, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        for _row in _response['DATA']:
                            self.DATA['CATEGORY_MAN_ID'].__setitem__(_row[1], _row[0])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '카테고리를 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {'THREAD': GetCategory(_thread_queue=self.APP.THREAD_QUEUE)}
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_category, True)
                    self.APP.POPUP.loading('카테고리를 불러오는 중 입니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetCategory',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '카테고리를 불러오는 중 오류가 발생하였습니다.')

    def request_get_lookup_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_lookup_list, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.WIDGETS['LOOKUP_TABLE'].delete_row()
                        self.WIDGETS['LOOKUP_TABLE'].insert_row(_response['DATA'])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '조회 내역을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetLookUpList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _lookup_date_start=self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'].get_start_value(),
                            _lookup_date_close=self.WIDGETS['LOOKUP_LC_LOOKUP_DATE'].get_close_value(),
                            _manager_index=self.WIDGETS['LOOKUP_LC_ID'].get_value_index()
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_lookup_list, True)
                    self.APP.POPUP.loading('조회 내역을 불러오는 중 입니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetLookUpList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '조회 내역을 불러오는 중 오류가 발생하였습니다.')

    def request_get_bl_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_bl_list, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.DATA['KEY_INDEX_LIST'].clear()
                        self.WIDGETS['BL_TABLE'].delete_row()
                        self.WIDGETS['BL_TABLE'].set_count_total(len(_response['DATA']))
                        self.WIDGETS['BL_TABLE'].insert_row(_response['DATA'])
                        self.DATA['KEY_INDEX_LIST'] = [_row[0] for _row in _response['DATA']]
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', 'B/L 번호 목록을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetBlList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _lookup_index=self.DATA['SELECTED_LOOKUP_INDEX'],
                            _order_str=self.WIDGETS['BL_TABLE'].get_head_sort()
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_bl_list, True)
                    self.APP.POPUP.loading('B/L 번호 목록을 불러오는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetBlList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', 'B/L 번호 목록을 불러오는 중 오류가 발생하였습니다.')

    def request_get_bl_detail_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_bl_detail_list, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.DATA['BL_DETAIL_DATA'] = _response['DATA']
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', 'B/L 번호 조회 결과를 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    _key_index = self.APP.APP_QUEUE.get()
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetBlDetailList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _app_queue=self.APP.APP_QUEUE,
                            _key_index=_key_index
                        ),
                        'ERROR_LIST': []
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_bl_detail_list, True)
                    self.APP.POPUP.loading('B/L 번호 조회 결과를 불러오는 중입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetBlDetailList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', 'B/L 번호 조회 결과를 불러오는 중 오류가 발생하였습니다.')

    def request_download_data_to_excel(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_download_data_to_excel, True)
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
                        messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': DownloadBlDataToExcel(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _download_path=self.getvar('MP') if self.getvar('MP') != '' else '.',
                            _excel_name='{}_{}.xlsx'.format(self.getvar('MID'), datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d_%H%M%S")),
                            _head=[_value['TEXT'] for _key, _value in self.APP.BL_HEAD_CATEGORY.items() if _value['IN_EXCEL']],
                            _lookup_index=self.DATA['SELECTED_LOOKUP_INDEX'],
                            _order_str=self.WIDGETS['BL_TABLE'].get_head_sort()
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_download_data_to_excel, True)
                    self.APP.POPUP.loading('엑셀 다운로드 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestDownloadDataToExcel',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')

    def request_download_data_to_print_form(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_download_data_to_print_form, True)
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
                        messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')
            else:
                pass
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': DownloadBlDataToPrintForm(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _download_path=self.getvar('MP') if self.getvar('MP') != '' else '.',
                            _excel_name='{}_{}.xlsx'.format(self.getvar('MID'), datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d_%H%M%S")),
                            _lookup_index=self.DATA['SELECTED_LOOKUP_INDEX']
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_download_data_to_print_form, True)
                    self.APP.POPUP.loading('엑셀 다운로드 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestDownloadDataToPrintForm',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 다운로드 중 오류가 발생하였습니다.')

    def request_bl_detail_print(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_bl_detail_print, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if not _response['SUCCESS']:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '인쇄용 엑셀을 생성 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    _printer_name = self.APP.APP_QUEUE.get()
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': PrintBlDetail(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _temp_path=self.APP.PATH['MASTER'],
                            _head_data=[self.DATA['BL_DETAIL_DATA'][0][0], self.DATA['BL_DETAIL_DATA'][0][1]],
                            _body_data = [_row[2:] for _row in self.DATA['BL_DETAIL_DATA']],
                            _printer_name=_printer_name
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_bl_detail_print, True)
                    self.APP.POPUP.loading('인쇄용 엑셀을 생성 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestBlDetailPrint',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '인쇄용 엑셀을 생성 중 오류가 발생하였습니다.')