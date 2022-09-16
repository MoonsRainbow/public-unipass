import os
import sys
import queue
import tkinter
import datetime
from tkinter import messagebox
from Model.model_vehicle_list import GetCategory, GetVehicleList
from Model.model_common import DownloadVehicleDataToExcel, GetLookUpDataDetail
from View import Body, LabelText, LabelEntry, LabelCombo, LabelCalendar, Table, Pagination


class VehicleListBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewVehicleList_VehicleListBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA['FILTER'] = {}
            self.DATA['CATEGORY_MAN_ID'] = {'All': 0}
            self.DATA['CATEGORY_VEHICLE_STATUS'] = {'All': 0}
            self.DATA['CATEGORY_INSPECTION'] = {'All': 0}
            self.DATA['CATEGORY_LIST_TYPE'] = {'최신 목록': 0, '누적 목록': 1}
            self.DATA['CATEGORY_VIEW_COUNT'] = {'100개': 100, '200개': 200, '500개': 500, '1000개': 1000, '2000개': 2000}
            self.WIDGETS.clear()
            self.request_get_category()
            self.WIDGETS['LE_EXPORT'] = LabelEntry(self, '수출신고번호')
            self.WIDGETS['LE_CHASSIS'] = LabelEntry(self, '차대번호')
            self.WIDGETS['LC_VEHICLE_STATUS'] = LabelCombo(self, '차량진행상태', self.DATA['CATEGORY_VEHICLE_STATUS'])
            self.WIDGETS['LC_INSPECTION'] = LabelCombo(self, '검사유무', self.DATA['CATEGORY_INSPECTION'])
            self.WIDGETS['LC_LOOKUP_DATE'] = LabelCalendar(self, '조회일자')
            self.WIDGETS['LC_REPORT_DATE'] = LabelCalendar(self, '신고일자', _is_on_off=True)
            self.WIDGETS['LC_LOAD_DUTY_DATE'] = LabelCalendar(self, '적재의무기한', _is_on_off=True)
            self.WIDGETS['LE_EXPORTER_SHIPPER'] = LabelEntry(self, '수출화주/대행자')
            self.WIDGETS['LC_ID'] = LabelCombo(self, '조회 아이디', self.DATA['CATEGORY_MAN_ID'])
            self.WIDGETS['LC_LIST_TYPE'] = LabelCombo(self, '목록 종류', self.DATA['CATEGORY_LIST_TYPE'])
            self.WIDGETS['TABLE'] = Table(self.APP, self)
            self.WIDGETS['TABLE'].set_table_resizing(True)
            self.WIDGETS['TABLE'].set_table_paging(True)
            self.WIDGETS['TABLE'].set_table_celling(True)
            self.WIDGETS['TABLE'].set_head_sort_command(_sort_command=self.event_table_head_click)
            self.WIDGETS['TABLE'].set_cell_double_click_command(_double_click_command=self.event_table_cell_double_click)
            self.WIDGETS['TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.VEHICLE_HEAD_CATEGORY.items() if 'VL' in _value['TABLE']},
                _is_change=True
            )
            self.WIDGETS['TABLE'].cell_clear()
            self.WIDGETS['TABLE'].clear_count()
            self.WIDGETS['TABLE'].delete_row()
            self.WIDGETS['LC_TOTAL'] = LabelText(self, '검색 결과', self.WIDGETS['TABLE'].COUNT['TOT'])
            self.WIDGETS['LC_VIEW_COUNT'] = LabelCombo(self, '행 개수', self.DATA['CATEGORY_VIEW_COUNT'])
            self.WIDGETS['PAGINATION'] = Pagination(self, self.WIDGETS['TABLE'].PAGE_NO, self.WIDGETS['TABLE'].END_PAGE, self.event_button_click_prev_page, self.event_button_click_move_page, self.event_button_click_next_page)
            self.WIDGETS['B_COPY'] = tkinter.Button(self, textvariable=self.WIDGETS['TABLE'].CELL_SELECTION_COUNT, overrelief=tkinter.SOLID, command=self.event_button_click_selection_copy, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_DOWNLOAD'] = tkinter.Button(self, text='엑셀\n다운로드', overrelief=tkinter.SOLID, command=self.event_button_click_excel_download, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_RESET'] = tkinter.Button(self, text='검색 필터\n초기화', overrelief=tkinter.SOLID, command=self.event_button_click_reset, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_FOLD'] = tkinter.Button(self, text='검색 필터\n가리기', overrelief=tkinter.SOLID, command=self.event_button_click_fold, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_UNFOLD'] = tkinter.Button(self, text='검색 필터\n보기', overrelief=tkinter.SOLID, command=self.event_button_click_unfold, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_SEARCH'] = tkinter.Button(self, text='검색', overrelief=tkinter.SOLID, command=self.event_button_click_search, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['TABLE'].set_location_y(284)
            self.WIDGETS['LE_EXPORT'].place(height=56, relwidth=0.24325)
            self.WIDGETS['LE_CHASSIS'].place(height=56, relwidth=0.24325, relx=0.25225)
            self.WIDGETS['LC_VEHICLE_STATUS'].place(height=56, relwidth=0.24325, relx=0.5045)
            self.WIDGETS['LC_INSPECTION'].place(height=56, relwidth=0.24325, relx=0.75675)
            self.WIDGETS['LC_LOOKUP_DATE'].place(height=56, relwidth=0.24325, y=66)
            self.WIDGETS['LC_REPORT_DATE'].place(height=56, relwidth=0.24325, relx=0.25225, y=66)
            self.WIDGETS['LC_LOAD_DUTY_DATE'].place(height=56, relwidth=0.24325, relx=0.5045, y=66)
            self.WIDGETS['LE_EXPORTER_SHIPPER'].place(height=56, relwidth=0.24325, relx=0.75675, y=66)
            self.WIDGETS['LC_ID'].place(height=56, relwidth=0.24325, y=132)
            self.WIDGETS['LC_LIST_TYPE'].place(height=56, relwidth=0.24325, relx=0.25225, y=132)
            self.WIDGETS['TABLE'].LOCATION.place(y=264, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 284)
            self.WIDGETS['LC_TOTAL'].place(height=56, relwidth=0.07, y=198)
            self.WIDGETS['LC_VIEW_COUNT'].place(height=56, relwidth=0.12, relx=0.079, y=198)
            self.WIDGETS['PAGINATION'].place(height=56, relwidth=0.2, relx=0.208, y=198)
            self.WIDGETS['B_COPY'].place(height=56, relwidth=0.06, relx=0.665, y=198)
            self.WIDGETS['B_DOWNLOAD'].place(height=56, relwidth=0.06, relx=0.734, y=198)
            self.WIDGETS['B_FOLD'].place(height=56, relwidth=0.06, relx=0.803, y=198)
            self.WIDGETS['B_RESET'].place(height=56, relwidth=0.06, relx=0.872, y=198)
            self.WIDGETS['B_SEARCH'].place(height=56, relwidth=0.06, relx=0.941, y=198)
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
        finally:
            self.WIDGETS['LC_VIEW_COUNT'].set_combo_selected_command(self.event_combo_select_view_count)
            self.APP.bind('<Control-c>', self.event_button_click_selection_copy)
            self.request_get_vehicle_list()

    def event_button_click_prev_page(self):
        if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_prev_page():
            self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_vehicle_list()

    def event_button_click_move_page(self):
        try:
            if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_move_page(self.WIDGETS['PAGINATION'].get_page_no()):
                self.WIDGETS['PAGINATION'].set_page_no()
                self.request_get_vehicle_list()
            else:
                if not self.WIDGETS['TABLE'].validate_page_no(self.WIDGETS['PAGINATION'].get_page_no()):
                    self.WIDGETS['PAGINATION'].set_page_no()
        except ValueError:
            messagebox.showwarning('UNIPASS ::', '숫자만 입력해주세요.')

    def event_button_click_next_page(self):
        if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_next_page():
            self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_vehicle_list()

    def event_button_click_selection_copy(self, _event=None):
        _tsc = None
        try:
            if len(self.WIDGETS['TABLE'].CELL_SELECTIONS) != 0:
                _tsc = self.WIDGETS['TABLE'].selection_copy()
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

    def event_button_click_reset(self):
        self.WIDGETS['LE_EXPORT'].clear_value()
        self.WIDGETS['LE_CHASSIS'].clear_value()
        self.WIDGETS['LC_VEHICLE_STATUS'].clear_value()
        self.WIDGETS['LC_INSPECTION'].clear_value()
        self.WIDGETS['LC_LOOKUP_DATE'].clear_value()
        self.WIDGETS['LC_REPORT_DATE'].clear_value()
        self.WIDGETS['LC_LOAD_DUTY_DATE'].clear_value()
        self.WIDGETS['LE_EXPORTER_SHIPPER'].clear_value()
        self.WIDGETS['LC_ID'].clear_value()
        self.WIDGETS['LC_LIST_TYPE'].clear_value()

    def event_button_click_fold(self):
        self.WIDGETS['LE_EXPORT'].place_forget()
        self.WIDGETS['LE_CHASSIS'].place_forget()
        self.WIDGETS['LC_VEHICLE_STATUS'].place_forget()
        self.WIDGETS['LC_INSPECTION'].place_forget()
        self.WIDGETS['LC_LOOKUP_DATE'].place_forget()
        self.WIDGETS['LC_REPORT_DATE'].place_forget()
        self.WIDGETS['LC_LOAD_DUTY_DATE'].place_forget()
        self.WIDGETS['LE_EXPORTER_SHIPPER'].place_forget()
        self.WIDGETS['LC_ID'].place_forget()
        self.WIDGETS['LC_LIST_TYPE'].place_forget()
        self.WIDGETS['TABLE'].set_location_y(86)
        self.WIDGETS['TABLE'].LOCATION.place_configure(y=66, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 86)
        self.WIDGETS['LC_TOTAL'].place_configure(height=56, relwidth=0.07, y=0)
        self.WIDGETS['LC_VIEW_COUNT'].place_configure(height=56, relwidth=0.12, relx=0.079, y=0)
        self.WIDGETS['PAGINATION'].place_configure(height=56, relwidth=0.2, relx=0.208, y=0)
        self.WIDGETS['B_COPY'].place_configure(height=56, relwidth=0.06, relx=0.665, y=0)
        self.WIDGETS['B_DOWNLOAD'].place_configure(height=56, relwidth=0.06, relx=0.734, y=0)
        self.WIDGETS['B_FOLD'].place_forget()
        self.WIDGETS['B_UNFOLD'].place(height=56, relwidth=0.06, relx=0.803, y=0)
        self.WIDGETS['B_RESET'].place_configure(height=56, relwidth=0.06, relx=0.872, y=0)
        self.WIDGETS['B_SEARCH'].place_configure(height=56, relwidth=0.06, relx=0.941, y=0)

    def event_button_click_unfold(self):
        self.WIDGETS['LE_EXPORT'].place(height=56, relwidth=0.24325)
        self.WIDGETS['LE_CHASSIS'].place(height=56, relwidth=0.24325, relx=0.25225)
        self.WIDGETS['LC_VEHICLE_STATUS'].place(height=56, relwidth=0.24325, relx=0.5045)
        self.WIDGETS['LC_INSPECTION'].place(height=56, relwidth=0.24325, relx=0.75675)
        self.WIDGETS['LC_LOOKUP_DATE'].place(height=56, relwidth=0.24325, y=66)
        self.WIDGETS['LC_REPORT_DATE'].place(height=56, relwidth=0.24325, relx=0.25225, y=66)
        self.WIDGETS['LC_LOAD_DUTY_DATE'].place(height=56, relwidth=0.24325, relx=0.5045, y=66)
        self.WIDGETS['LE_EXPORTER_SHIPPER'].place(height=56, relwidth=0.24325, relx=0.75675, y=66)
        self.WIDGETS['LC_ID'].place(height=56, relwidth=0.24325, y=132)
        self.WIDGETS['LC_LIST_TYPE'].place(height=56, relwidth=0.24325, relx=0.25225, y=132)
        self.WIDGETS['TABLE'].set_location_y(284)
        self.WIDGETS['TABLE'].LOCATION.place_configure(y=264, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 284)
        self.WIDGETS['LC_TOTAL'].place_configure(height=56, relwidth=0.07, y=198)
        self.WIDGETS['LC_VIEW_COUNT'].place_configure(height=56, relwidth=0.12, relx=0.079, y=198)
        self.WIDGETS['PAGINATION'].place_configure(height=56, relwidth=0.2, relx=0.208, y=198)
        self.WIDGETS['B_COPY'].place_configure(height=56, relwidth=0.06, relx=0.665, y=198)
        self.WIDGETS['B_DOWNLOAD'].place_configure(height=56, relwidth=0.06, relx=0.734, y=198)
        self.WIDGETS['B_FOLD'].place(height=56, relwidth=0.06, relx=0.803, y=198)
        self.WIDGETS['B_UNFOLD'].place_forget()
        self.WIDGETS['B_RESET'].place_configure(height=56, relwidth=0.06, relx=0.872, y=198)
        self.WIDGETS['B_SEARCH'].place_configure(height=56, relwidth=0.06, relx=0.941, y=198)

    def event_button_click_search(self):
        self.WIDGETS['TABLE'].head_sort_clear()
        self.WIDGETS['TABLE'].set_move_page(1)
        self.WIDGETS['PAGINATION'].set_page_no()
        self.request_get_vehicle_list()

    def event_combo_select_view_count(self, _event=None):
        try:
            if not self.WIDGETS['TABLE'].validate_end_page(self.WIDGETS['LC_VIEW_COUNT'].get_value_index()):
                self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_vehicle_list()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventComboSelectViewCount',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '행 개수 변경 중 오류가 발생하였습니다.')

    def event_table_head_click(self, _head_key: str = ''):
        try:
            self.WIDGETS['TABLE'].head_sort(_head_key)
            self.request_get_vehicle_list()
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
            if not self.APP.POPUP.winfo_exists():
                self.APP.APP_QUEUE.put(_ici)
                self.request_get_lookup_data_detail()
                self.APP.POPUP.detail_data(self.APP.APP_QUEUE.get(), self.event_popup_cell_click)
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
                        for _row in _response['DATA'][0]:
                            self.DATA['CATEGORY_MAN_ID'].__setitem__(_row[1], _row[0])
                        for _row_no in range(len(_response['DATA'][1])):
                            self.DATA['CATEGORY_VEHICLE_STATUS'].__setitem__(_response['DATA'][1][_row_no][0], _row_no)
                        for _row_no in range(len(_response['DATA'][2])):
                            self.DATA['CATEGORY_INSPECTION'].__setitem__(_response['DATA'][2][_row_no][0], _row_no)
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

    def request_get_vehicle_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_vehicle_list, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.WIDGETS['TABLE'].delete_row()
                        self.WIDGETS['TABLE'].set_count_total(_response['DATA'][0][0][0])
                        self.WIDGETS['TABLE'].set_end_page(_response['DATA'][0][0][1])
                        for _row_no in range(len(_response['DATA'][1])):
                            _row_values = list(_response['DATA'][1][_row_no])
                            for _col_index in range(len(_row_values)):
                                if _row_values[_col_index] is None:
                                    _row_values[_col_index] = ''
                            if type(_row_values[5]) == datetime.date:
                                _rdd_d_day = (_row_values[5] - datetime.date.today()).days
                                if _rdd_d_day < 0:
                                    _row_values[5] = 'OVER) {}'.format(_row_values[5])
                                elif _rdd_d_day <= 10:
                                    _row_values[5] = 'D-{}) {}'.format(str(_rdd_d_day).zfill(2), _row_values[5])
                            else:
                                _rdd_d_day = 999999999
                            if _rdd_d_day <= 10:
                                self.WIDGETS['TABLE'].insert(parent='', values=_row_values, index=_row_no, tags='rdd_warn')
                            else:
                                self.WIDGETS['TABLE'].insert(parent='', values=_row_values, index=_row_no)
                        self.WIDGETS['TABLE'].tag_configure("rdd_warn", foreground='red')
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '차량 목록을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.WIDGETS['TABLE'].cell_clear()
                    _en_list = []
                    for _en in str(self.WIDGETS['LE_EXPORT'].get_value()).split(' '):
                        _ = []
                        if '\t\n' in _en:
                            _ += _en.split('\t\n')
                        elif '\t' in _en:
                            _ += _en.split('\t')
                        elif '\n' in _en:
                            _ += _en.split('\n')
                        else:
                            _.append(_en)
                        for __ in _:
                            if len(__) > 1:
                                _en_list.append(str(__).strip().upper())
                    _cn_list = []
                    for _cn in str(self.WIDGETS['LE_CHASSIS'].get_value()).split(' '):
                        _ = []
                        if '\t\n' in _cn:
                            _ += _cn.split('\t\n')
                        elif '\t' in _cn:
                            _ += _cn.split('\t')
                        elif '\n' in _cn:
                            _ += _cn.split('\n')
                        else:
                            _.append(_cn)
                        for __ in _:
                            if len(__) > 1:
                                _cn_list.append(str(__).strip().upper())
                    _es = self.WIDGETS['LE_EXPORTER_SHIPPER'].get_value()
                    if '\t\n' in _es:
                        _es = _es.replace('\t\n', '')
                    if '\t' in _es:
                        _es = _es.replace('\t', '')
                    if '\n' in _es:
                        _es = _es.replace('\n', '')
                    self.DATA['FILTER'] = {
                        'EN_LIST': _en_list,
                        'CN_LIST': _cn_list,
                        'VEHICLE_STATUS': self.WIDGETS['LC_VEHICLE_STATUS'].get_value_text(),
                        'INSPECTION_FLAG': self.WIDGETS['LC_INSPECTION'].get_value_text(),
                        'LOOKUP_DATE_START': self.WIDGETS['LC_LOOKUP_DATE'].get_start_value(),
                        'LOOKUP_DATE_CLOSE': self.WIDGETS['LC_LOOKUP_DATE'].get_close_value(),
                        'REPORT_DATE_START': self.WIDGETS['LC_REPORT_DATE'].get_start_value(),
                        'REPORT_DATE_CLOSE': self.WIDGETS['LC_REPORT_DATE'].get_close_value(),
                        'LOAD_DUTY_DATE_START': self.WIDGETS['LC_LOAD_DUTY_DATE'].get_start_value(),
                        'LOAD_DUTY_DATE_CLOSE': self.WIDGETS['LC_LOAD_DUTY_DATE'].get_close_value(),
                        'EXPORTER_SHIPPER': _es,
                        'MANAGER_INDEX': self.WIDGETS['LC_ID'].get_value_index(),
                        'LIST_TYPE': self.WIDGETS['LC_LIST_TYPE'].get_value_index()
                    }
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetVehicleList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _en_list=self.DATA['FILTER']['EN_LIST'],
                            _cn_list=self.DATA['FILTER']['CN_LIST'],
                            _vehicle_status=self.DATA['FILTER']['VEHICLE_STATUS'],
                            _inspection_flag=self.DATA['FILTER']['INSPECTION_FLAG'],
                            _lookup_date_start=self.DATA['FILTER']['LOOKUP_DATE_START'],
                            _lookup_date_close=self.DATA['FILTER']['LOOKUP_DATE_CLOSE'],
                            _report_date_start=self.DATA['FILTER']['REPORT_DATE_START'],
                            _report_date_close=self.DATA['FILTER']['REPORT_DATE_CLOSE'],
                            _load_duty_date_start=self.DATA['FILTER']['LOAD_DUTY_DATE_START'],
                            _load_duty_date_close=self.DATA['FILTER']['LOAD_DUTY_DATE_CLOSE'],
                            _exporter_shipper_name=self.DATA['FILTER']['EXPORTER_SHIPPER'],
                            _manager_index=self.DATA['FILTER']['MANAGER_INDEX'],
                            _list_type=self.DATA['FILTER']['LIST_TYPE'],
                            _view_count=self.WIDGETS['LC_VIEW_COUNT'].get_value_index(),
                            _page_no=int(self.WIDGETS['TABLE'].get_now_page()),
                            _order_str=self.WIDGETS['TABLE'].get_head_sort(),
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_vehicle_list, True)
                    self.APP.POPUP.loading('차량 목록을 불러오는 중 입니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetVehicleList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '차량 목록을 불러오는 중 오류가 발생하였습니다.')

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
                        'THREAD': DownloadVehicleDataToExcel(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _download_path=self.getvar('MP') if self.getvar('MP') != '' else '.',
                            _excel_name='{}_{}.xlsx'.format(self.getvar('MID'), datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d_%H%M%S")),
                            _head=[_value['TEXT'] for _key, _value in self.APP.VEHICLE_HEAD_CATEGORY.items() if _value['IN_EXCEL']],
                            _en_list=self.DATA['FILTER']['EN_LIST'],
                            _cn_list=self.DATA['FILTER']['CN_LIST'],
                            _vehicle_status=self.DATA['FILTER']['VEHICLE_STATUS'],
                            _inspection_flag=self.DATA['FILTER']['INSPECTION_FLAG'],
                            _lookup_date_start=self.DATA['FILTER']['LOOKUP_DATE_START'],
                            _lookup_date_close=self.DATA['FILTER']['LOOKUP_DATE_CLOSE'],
                            _report_date_start=self.DATA['FILTER']['REPORT_DATE_START'],
                            _report_date_close=self.DATA['FILTER']['REPORT_DATE_CLOSE'],
                            _load_duty_date_start=self.DATA['FILTER']['LOAD_DUTY_DATE_START'],
                            _load_duty_date_close=self.DATA['FILTER']['LOAD_DUTY_DATE_CLOSE'],
                            _exporter_shipper_name=self.DATA['FILTER']['EXPORTER_SHIPPER'],
                            _manager_index=self.DATA['FILTER']['MANAGER_INDEX'],
                            _list_type=self.DATA['FILTER']['LIST_TYPE']
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

    def request_get_lookup_data_detail(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_lookup_data_detail, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    if _response['SUCCESS']:
                        self.APP.APP_QUEUE.put(_response['DATA'])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '상세 정보를 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    _ici = self.APP.APP_QUEUE.get()
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetLookUpDataDetail(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_index=self.WIDGETS['LC_ID'].get_value_index(),
                            _info_core_index=_ici
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_lookup_data_detail, True)
                    self.APP.POPUP.loading('상세 정보를 불러오는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetLookUpDataDetail',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '상세 정보를 불러오는 중 오류가 발생하였습니다.')
