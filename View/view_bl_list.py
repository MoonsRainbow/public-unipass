import os
import sys
import xlrd
import queue
import tkinter
import datetime
from tkinter import messagebox, filedialog
from Model.model_common import DownloadBlDataToExcel, PrintBlDetail
from Model.model_bl_list import GetCategory, GetBlList, LookUpDataFromUnipass, GetBlDetailList
from View import Body, LabelEntryButton, LabelText, LabelEntry, LabelCombo, LabelCalendar, Table, Pagination


class BlListBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewBlList_BlListBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA['FILTER'] = {}
            self.DATA['CATEGORY_MAN_ID'] = {'All': 0}
            self.DATA['CATEGORY_SHIPMENT_PLACE'] = {'All': 0}
            self.DATA['CATEGORY_PRE_SHIPPING'] = {'All': 0}
            self.DATA['CATEGORY_LIST_TYPE'] = {'최신 목록': 0, '누적 목록': 1}
            self.DATA['CATEGORY_VIEW_COUNT'] = {'100개': 100, '200개': 200, '500개': 500, '1000개': 1000, '2000개': 2000}
            self.DATA['CATEGORY_COLUMN_TITLE'] = []
            self.DATA['BL_LIST'] = []
            self.DATA['BL_DETAIL_DATA'] = None
            self.DATA['KEY_INDEX_LIST'] = []
            self.WIDGETS.clear()
            self.request_get_category()
            self.WIDGETS['LEB_BL_NUMBER'] = LabelEntryButton(self, 'B/L 번호', '등록', self.event_button_click_bl_lookup)
            self.WIDGETS['B_EXCEL'] = tkinter.Button(self, text='엑셀 선택', overrelief=tkinter.SOLID, command=self.event_button_click_excel_choice, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['LE_BL_NUMBER'] = LabelEntry(self, 'B/L 번호')
            self.WIDGETS['LE_EXPORT_NUMBER'] = LabelEntry(self, '수출신고번호')
            self.WIDGETS['LE_MANAGEMENT_NUMBER'] = LabelEntry(self, '적하목록관리번호')
            self.WIDGETS['LE_EXPORT_SHIPPER'] = LabelEntry(self, '수출자')
            self.WIDGETS['LC_LOOKUP_DATE'] = LabelCalendar(self, '조회일자')
            self.WIDGETS['LC_ACCEPT_DATE'] = LabelCalendar(self, '수리일자', _is_on_off=True)
            self.WIDGETS['LC_LOAD_DUTY_DEADLINE'] = LabelCalendar(self, '적재의무기한', _is_on_off=True)
            self.WIDGETS['LC_DEPARTURE_DATE'] = LabelCalendar(self, '출항일자', _is_on_off=True)
            self.WIDGETS['LC_SHIPMENT_PLACE'] = LabelCombo(self, '선기적지', self.DATA['CATEGORY_SHIPMENT_PLACE'])
            self.WIDGETS['LC_PRE_SHIPPING'] = LabelCombo(self, '선기적완료여부', self.DATA['CATEGORY_PRE_SHIPPING'])
            self.WIDGETS['LC_ID'] = LabelCombo(self, '조회 아이디', self.DATA['CATEGORY_MAN_ID'])
            self.WIDGETS['LC_LIST_TYPE'] = LabelCombo(self, '목록 종류', self.DATA['CATEGORY_LIST_TYPE'])
            self.WIDGETS['TABLE'] = Table(self.APP, self)
            self.WIDGETS['TABLE'].set_table_resizing(True)
            self.WIDGETS['TABLE'].set_table_paging(True)
            self.WIDGETS['TABLE'].set_table_celling(True)
            self.WIDGETS['TABLE'].set_head_sort_command(_sort_command=self.event_table_head_click)
            self.WIDGETS['TABLE'].set_cell_double_click_command(_double_click_command=self.event_table_cell_double_click)
            self.WIDGETS['TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.BL_HEAD_CATEGORY.items() if 'BL' in _value['TABLE']},
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
            self.WIDGETS['TABLE'].set_location_y(350)
            self.WIDGETS['LEB_BL_NUMBER'].place(relwidth=0.871, height=56)
            self.WIDGETS['B_EXCEL'].place(relx=0.88, relwidth=0.12, height=56)
            self.WIDGETS['LE_BL_NUMBER'].place(height=56, y=66, relwidth=0.24325)
            self.WIDGETS['LE_EXPORT_NUMBER'].place(height=56, y=66, relwidth=0.24325, relx=0.25225)
            self.WIDGETS['LE_MANAGEMENT_NUMBER'].place(height=56, y=66, relwidth=0.24325, relx=0.5045)
            self.WIDGETS['LE_EXPORT_SHIPPER'].place(height=56, y=66, relwidth=0.24325, relx=0.75675)
            self.WIDGETS['LC_LOOKUP_DATE'].place(height=56, relwidth=0.24325, y=132)
            self.WIDGETS['LC_ACCEPT_DATE'].place(height=56, relwidth=0.24325, relx=0.25225, y=132)
            self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].place(height=56, relwidth=0.24325, relx=0.5045, y=132)
            self.WIDGETS['LC_DEPARTURE_DATE'].place(height=56, relwidth=0.24325, relx=0.75675, y=132)
            self.WIDGETS['LC_SHIPMENT_PLACE'].place(height=56, relwidth=0.24325, y=198)
            self.WIDGETS['LC_PRE_SHIPPING'].place(height=56, relwidth=0.24325, relx=0.25225, y=198)
            self.WIDGETS['LC_ID'].place(height=56, relwidth=0.24325, relx=0.5045, y=198)
            self.WIDGETS['LC_LIST_TYPE'].place(height=56, relwidth=0.24325, relx=0.75675, y=198)
            self.WIDGETS['TABLE'].LOCATION.place(y=330, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 350)
            self.WIDGETS['LC_TOTAL'].place(height=56, relwidth=0.07, y=264)
            self.WIDGETS['LC_VIEW_COUNT'].place(height=56, relwidth=0.12, relx=0.079, y=264)
            self.WIDGETS['PAGINATION'].place(height=56, relwidth=0.2, relx=0.208, y=264)
            self.WIDGETS['B_COPY'].place(height=56, relwidth=0.06, relx=0.665, y=264)
            self.WIDGETS['B_DOWNLOAD'].place(height=56, relwidth=0.06, relx=0.734, y=264)
            self.WIDGETS['B_FOLD'].place(height=56, relwidth=0.06, relx=0.803, y=264)
            self.WIDGETS['B_RESET'].place(height=56, relwidth=0.06, relx=0.872, y=264)
            self.WIDGETS['B_SEARCH'].place(height=56, relwidth=0.06, relx=0.941, y=264)
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
            self.request_get_bl_list()

    def event_button_click_bl_lookup(self):
        try:
            if len(self.WIDGETS['LEB_BL_NUMBER'].get_value()) != 0:
                _inserted_bl_list = str(self.WIDGETS['LEB_BL_NUMBER'].get_value()).split(' ')
                _bl_list = []
                for _bl in _inserted_bl_list:
                    if '\t\n' in _bl:
                        _bl_list += _bl.split('\t\n')
                    elif '\t' in _bl:
                        _bl_list += _bl.split('\t')
                    elif '\n' in _bl:
                        _bl_list += _bl.split('\n')
                    else:
                        _bl_list.append(_bl)
                for _bl in _bl_list:
                    _bl = _bl.strip().upper()
                    if len(_bl) > 1:
                        if _bl not in [_[0] for _ in self.DATA['BL_LIST']]:
                            self.DATA['BL_LIST'].append([_bl, 0])
                if len(self.DATA['BL_LIST']) > 0:
                    self.request_lookup_data_from_unipass()
                    self.request_get_bl_list()
            self.WIDGETS['LEB_BL_NUMBER'].set_value()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickBlLookUp',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '입력된 B/L 번호를 읽어오는 중 오류가 발생하였습니다.')

    def event_button_click_excel_choice(self):
        try:
            _excel_path = self.request_choice_excel()
            if _excel_path != '':
                _cer = self.request_choice_sheet(_excel_path)
                if _cer[0]:
                    if _cer[1] is not None:
                        _rsr = self.request_read_sheet(_excel_path, _cer[1])
                        if _rsr[0]:
                            self.DATA['BL_LIST'].clear()
                            self.DATA['BL_LIST'] = _rsr[1]
                            self.request_lookup_data_from_unipass()
                            self.request_get_bl_list()
                        else:
                            if _rsr[1] is None:
                                messagebox.showwarning('UNIPASS ::', '\'{}\' 시트에서\nB/L 번호를 찾을 수 없습니다.\n확인 후 다시 시도해주세요.'.format(_cer[1]))
                            else:
                                self.APP.request_error_report_db(_error_info=_rsr[1])
                                messagebox.showerror('UNIPASS :: ', '엑셀 시트 읽기 중 오류가 발생하였습니다.')
                else:
                    self.APP.request_error_report_db(_error_info=_cer[1])
                    messagebox.showerror('UNIPASS :: ', '엑셀 시트 선택 중 오류가 발생하였습니다.')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickExcelChoice',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 / 시트 선택 중 오류가 발생하였습니다.')

    def event_button_click_prev_page(self):
        if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_prev_page():
            self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_bl_list()

    def event_button_click_move_page(self):
        try:
            if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_move_page(self.WIDGETS['PAGINATION'].get_page_no()):
                self.WIDGETS['PAGINATION'].set_page_no()
                self.request_get_bl_list()
            else:
                if not self.WIDGETS['TABLE'].validate_page_no(self.WIDGETS['PAGINATION'].get_page_no()):
                    self.WIDGETS['PAGINATION'].set_page_no()
        except ValueError:
            messagebox.showwarning('UNIPASS ::', '숫자만 입력해주세요.')

    def event_button_click_next_page(self):
        if self.WIDGETS['TABLE'].get_now_page() != self.WIDGETS['TABLE'].set_next_page():
            self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_bl_list()

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
        self.WIDGETS['LE_BL_NUMBER'].clear_value()
        self.WIDGETS['LE_EXPORT_NUMBER'].clear_value()
        self.WIDGETS['LE_MANAGEMENT_NUMBER'].clear_value()
        self.WIDGETS['LE_EXPORT_SHIPPER'].clear_value()
        self.WIDGETS['LC_LOOKUP_DATE'].clear_value()
        self.WIDGETS['LC_ACCEPT_DATE'].clear_value()
        self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].clear_value()
        self.WIDGETS['LC_DEPARTURE_DATE'].clear_value()
        self.WIDGETS['LC_SHIPMENT_PLACE'].clear_value()
        self.WIDGETS['LC_PRE_SHIPPING'].clear_value()
        self.WIDGETS['LC_ID'].clear_value()
        self.WIDGETS['LC_LIST_TYPE'].clear_value()

    def event_button_click_fold(self):
        self.WIDGETS['LEB_BL_NUMBER'].place_forget()
        self.WIDGETS['B_EXCEL'].place_forget()
        self.WIDGETS['LE_BL_NUMBER'].place_forget()
        self.WIDGETS['LE_EXPORT_NUMBER'].place_forget()
        self.WIDGETS['LE_MANAGEMENT_NUMBER'].place_forget()
        self.WIDGETS['LE_EXPORT_SHIPPER'].place_forget()
        self.WIDGETS['LC_LOOKUP_DATE'].place_forget()
        self.WIDGETS['LC_ACCEPT_DATE'].place_forget()
        self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].place_forget()
        self.WIDGETS['LC_DEPARTURE_DATE'].place_forget()
        self.WIDGETS['LC_SHIPMENT_PLACE'].place_forget()
        self.WIDGETS['LC_PRE_SHIPPING'].place_forget()
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
        self.WIDGETS['LEB_BL_NUMBER'].place(relwidth=0.871, height=56)
        self.WIDGETS['B_EXCEL'].place(relx=0.88, relwidth=0.12, height=56)
        self.WIDGETS['LE_BL_NUMBER'].place(height=56, y=66, relwidth=0.24325)
        self.WIDGETS['LE_EXPORT_NUMBER'].place(height=56, y=66, relwidth=0.24325, relx=0.25225)
        self.WIDGETS['LE_MANAGEMENT_NUMBER'].place(height=56, y=66, relwidth=0.24325, relx=0.5045)
        self.WIDGETS['LE_EXPORT_SHIPPER'].place(height=56, y=66, relwidth=0.24325, relx=0.75675)
        self.WIDGETS['LC_LOOKUP_DATE'].place(height=56, relwidth=0.24325, y=132)
        self.WIDGETS['LC_ACCEPT_DATE'].place(height=56, relwidth=0.24325, relx=0.25225, y=132)
        self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].place(height=56, relwidth=0.24325, relx=0.5045, y=132)
        self.WIDGETS['LC_DEPARTURE_DATE'].place(height=56, relwidth=0.24325, relx=0.75675, y=132)
        self.WIDGETS['LC_SHIPMENT_PLACE'].place(height=56, relwidth=0.24325, y=198)
        self.WIDGETS['LC_PRE_SHIPPING'].place(height=56, relwidth=0.24325, relx=0.25225, y=198)
        self.WIDGETS['LC_ID'].place(height=56, relwidth=0.24325, relx=0.5045, y=198)
        self.WIDGETS['LC_LIST_TYPE'].place(height=56, relwidth=0.24325, relx=0.75675, y=198)
        self.WIDGETS['TABLE'].set_location_y(350)
        self.WIDGETS['TABLE'].LOCATION.place_configure(y=330, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 350)
        self.WIDGETS['LC_TOTAL'].place_configure(height=56, relwidth=0.07, y=264)
        self.WIDGETS['LC_VIEW_COUNT'].place_configure(height=56, relwidth=0.12, relx=0.079, y=264)
        self.WIDGETS['PAGINATION'].place_configure(height=56, relwidth=0.2, relx=0.208, y=264)
        self.WIDGETS['B_COPY'].place_configure(height=56, relwidth=0.06, relx=0.665, y=264)
        self.WIDGETS['B_DOWNLOAD'].place_configure(height=56, relwidth=0.06, relx=0.734, y=264)
        self.WIDGETS['B_FOLD'].place(height=56, relwidth=0.06, relx=0.803, y=264)
        self.WIDGETS['B_UNFOLD'].place_forget()
        self.WIDGETS['B_RESET'].place_configure(height=56, relwidth=0.06, relx=0.872, y=264)
        self.WIDGETS['B_SEARCH'].place_configure(height=56, relwidth=0.06, relx=0.941, y=264)

    def event_button_click_search(self):
        self.WIDGETS['TABLE'].head_sort_clear()
        self.WIDGETS['TABLE'].set_move_page(1)
        self.WIDGETS['PAGINATION'].set_page_no()
        self.request_get_bl_list()

    def event_combo_select_view_count(self, _event=None):
        try:
            if not self.WIDGETS['TABLE'].validate_end_page(self.WIDGETS['LC_VIEW_COUNT'].get_value_index()):
                self.WIDGETS['PAGINATION'].set_page_no()
            self.request_get_bl_list()
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
                        for _row in _response['DATA'][0]:
                            self.DATA['CATEGORY_MAN_ID'].__setitem__(_row[1], _row[0])
                        for _row_no in range(len(_response['DATA'][1])):
                            self.DATA['CATEGORY_SHIPMENT_PLACE'].__setitem__(_response['DATA'][1][_row_no][0], _row_no)
                        for _row_no in range(len(_response['DATA'][2])):
                            self.DATA['CATEGORY_PRE_SHIPPING'].__setitem__(_response['DATA'][2][_row_no][0], _row_no)
                        for _row in _response['DATA'][3]:
                            self.DATA['CATEGORY_COLUMN_TITLE'].append(_row[0])
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
                        self.WIDGETS['TABLE'].delete_row()
                        self.WIDGETS['TABLE'].set_count_total(_response['DATA'][0][0][0])
                        self.WIDGETS['TABLE'].set_end_page(_response['DATA'][0][0][1])
                        self.WIDGETS['TABLE'].insert_row(_response['DATA'][1])
                        self.DATA['KEY_INDEX_LIST'] = [_row[0] for _row in _response['DATA'][1]]
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', 'B/L 목록을 불러오는 중 오류가 발생하였습니다.')
            else:
                pass
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.WIDGETS['TABLE'].cell_clear()
                    _bn_list = []
                    for _bn in str(self.WIDGETS['LE_BL_NUMBER'].get_value()).split(' '):
                        _ = []
                        if '\t\n' in _bn:
                            _ += _bn.split('\t\n')
                        elif '\t' in _bn:
                            _ += _bn.split('\t')
                        elif '\n' in _bn:
                            _ += _bn.split('\n')
                        else:
                            _.append(_bn)
                        for __ in _:
                            if len(__) > 1:
                                _bn_list.append(str(__).strip().upper())
                    _en_list = []
                    for _en in str(self.WIDGETS['LE_EXPORT_NUMBER'].get_value()).split(' '):
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
                    _mn_list = []
                    for _mn in str(self.WIDGETS['LE_MANAGEMENT_NUMBER'].get_value()).split(' '):
                        _ = []
                        if '\t\n' in _mn:
                            _ += _mn.split('\t\n')
                        elif '\t' in _mn:
                            _ += _mn.split('\t')
                        elif '\n' in _mn:
                            _ += _mn.split('\n')
                        else:
                            _.append(_mn)
                        for __ in _:
                            if len(__) > 1:
                                _mn_list.append(str(__).strip().upper())
                    _es = self.WIDGETS['LE_EXPORT_SHIPPER'].get_value()
                    if '\t\n' in _es:
                        _es = _es.replace('\t\n', '')
                    if '\t' in _es:
                        _es = _es.replace('\t', '')
                    if '\n' in _es:
                        _es = _es.replace('\n', '')
                    self.DATA['FILTER'] = {
                        'BN_LIST': _bn_list,
                        'EN_LIST': _en_list,
                        'MN_LIST': _mn_list,
                        'EXPORTER_SHIPPER': _es,
                        'LOOKUP_DATE_START': self.WIDGETS['LC_LOOKUP_DATE'].get_start_value(),
                        'LOOKUP_DATE_CLOSE': self.WIDGETS['LC_LOOKUP_DATE'].get_close_value(),
                        'ACCEPT_DATE_START': self.WIDGETS['LC_ACCEPT_DATE'].get_start_value(),
                        'ACCEPT_DATE_CLOSE': self.WIDGETS['LC_ACCEPT_DATE'].get_close_value(),
                        'LOAD_DUTY_DATE_START': self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].get_start_value(),
                        'LOAD_DUTY_DATE_CLOSE': self.WIDGETS['LC_LOAD_DUTY_DEADLINE'].get_close_value(),
                        'DEPARTURE_DATE_START': self.WIDGETS['LC_DEPARTURE_DATE'].get_start_value(),
                        'DEPARTURE_DATE_CLOSE': self.WIDGETS['LC_DEPARTURE_DATE'].get_close_value(),
                        'SHIPMENT_PLACE': self.WIDGETS['LC_SHIPMENT_PLACE'].get_value_text(),
                        'PRE_SHIPPING_FLAG': self.WIDGETS['LC_PRE_SHIPPING'].get_value_text(),
                        'MANAGER_INDEX': self.WIDGETS['LC_ID'].get_value_index(),
                        'LIST_TYPE': self.WIDGETS['LC_LIST_TYPE'].get_value_index()
                    }
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetBlList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _bn_list=self.DATA['FILTER']['BN_LIST'],
                            _en_list=self.DATA['FILTER']['EN_LIST'],
                            _mn_list=self.DATA['FILTER']['MN_LIST'],
                            _exporter_shipper_name=self.DATA['FILTER']['EXPORTER_SHIPPER'],
                            _lookup_date_start=self.DATA['FILTER']['LOOKUP_DATE_START'],
                            _lookup_date_close=self.DATA['FILTER']['LOOKUP_DATE_CLOSE'],
                            _accept_date_start=self.DATA['FILTER']['ACCEPT_DATE_START'],
                            _accept_date_close=self.DATA['FILTER']['ACCEPT_DATE_CLOSE'],
                            _load_duty_date_start=self.DATA['FILTER']['LOAD_DUTY_DATE_START'],
                            _load_duty_date_close=self.DATA['FILTER']['LOAD_DUTY_DATE_CLOSE'],
                            _departure_date_start=self.DATA['FILTER']['DEPARTURE_DATE_START'],
                            _departure_date_close=self.DATA['FILTER']['DEPARTURE_DATE_CLOSE'],
                            _shipment_place=self.DATA['FILTER']['SHIPMENT_PLACE'],
                            _pre_shipping_flag=self.DATA['FILTER']['PRE_SHIPPING_FLAG'],
                            _manager_index=self.DATA['FILTER']['MANAGER_INDEX'],
                            _list_type=self.DATA['FILTER']['LIST_TYPE'],
                            _view_count=self.WIDGETS['LC_VIEW_COUNT'].get_value_index(),
                            _page_no=int(self.WIDGETS['TABLE'].get_now_page()),
                            _order_str=self.WIDGETS['TABLE'].get_head_sort(),
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_bl_list, True)
                    self.APP.POPUP.loading('B/L 목록을 불러오는 중 입니다.')
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
            messagebox.showerror('UNIPASS :: ', 'B/L 목록을 불러오는 중 오류가 발생하였습니다.')

    def request_choice_excel(self):
        try:
            _chosen_excel_path = ''
            _chosen_excel_path = filedialog.askopenfilename(
                initialdir='./',
                title='엑셀 파일 선택',
                filetypes=[('Excel 통합문서', '*.xlsx')],
                defaultextension=''
            )
            if _chosen_excel_path != '':
                _selected_excel_path = str(_chosen_excel_path).replace('\\', '/')
            return _chosen_excel_path
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestChoiceExcel',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '엑셀 선택 중 오류가 발생하였습니다.')
            return ''

    def request_choice_sheet(self, _excel_path: str):
        _excel: [xlrd.open_workbook, None] = None
        try:
            _excel = xlrd.open_workbook(_excel_path)
            _sheet_list = _excel.sheet_names()
            self.APP.clear_queues()
            self.APP.POPUP.sheet_choice(_sheet_list)
            if self.APP.THREAD_QUEUE.empty():
                return True, None
            else:
                return True, self.APP.THREAD_QUEUE.get()
        except queue.Empty:
            pass
        except:
            return False, {
                'ECD': self._code + '_RequestChoiceSheet',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': [_excel_path]
            }
        finally:
            if type(_excel) is not None:
                _excel.release_resources()

    def request_read_sheet(self, _excel_path: str, _sheet_name):
        _excel: [xlrd.open_workbook, None] = None
        try:
            _excel = xlrd.open_workbook(_excel_path)
            _sheet = _excel.sheet_by_name(_sheet_name)
            _s_row, _s_col = None, None
            for _title in self.DATA['CATEGORY_COLUMN_TITLE']:
                for _row in range(0, _sheet.nrows if _sheet.nrows < 15 else 15):
                    for _col in range(0, _sheet.ncols if _sheet.ncols < 15 else 15):
                        if _title in str(_sheet.cell_value(_row, _col)):
                            _s_row, _s_col = _row, _col
                            break
                    if _s_row is not None and _s_col is not None:
                        break
                if _s_row is not None and _s_col is not None:
                    break
            if _s_row is not None and _s_col is not None:
                _bl_list = []
                for _row in range(_s_row + 1, _sheet.nrows):
                    _bl = str(_sheet.cell_value(_row, _s_col)).strip().upper()
                    if len(_bl) > 1:
                        if _bl not in [_[0] for _ in _bl_list]:
                            _bl_list.append([_bl, _row + 1])
                return True, _bl_list
            else:
                return False, None
        except:
            return False, {
                'ECD': self._code + '_RequestReadExcel',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': [_excel_path]
            }
        finally:
            if type(_excel) is not None:
                _excel.release_resources()

    def request_lookup_data_from_unipass(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_lookup_data_from_unipass, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['RUN']:
                        if _response['SUCCESS']:
                            self.APP.POPUP.DATA['SUCCESS'].set(str(int(self.APP.POPUP.DATA['SUCCESS'].get()) + 1))
                        else:
                            self.APP.CURRENT_THREAD_INFO['ERROR_LIST'].append(dict(_response['DATA']).copy())
                            self.APP.POPUP.DATA['FAIL'].set(str(int(self.APP.POPUP.DATA['FAIL'].get()) + 1))
                        self.APP.after(100, self.request_lookup_data_from_unipass, True)
                    else:
                        if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                            self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                        self.APP.POPUP.destroy()
                        for _error in self.APP.CURRENT_THREAD_INFO['ERROR_LIST']:
                            self.APP.request_error_report_db(_error_info=_error)
                        self.APP.CURRENT_THREAD_INFO.clear()
                        self.DATA['BL_LIST'].clear()
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': LookUpDataFromUnipass(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _app_queue=self.APP.APP_QUEUE,
                            _manager_index=int(self.APP.getvar('MI')),
                            _bl_list=self.DATA['BL_LIST']
                        ),
                        'ERROR_LIST': []
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_lookup_data_from_unipass, True)
                    self.APP.POPUP.lookup_data('유니패스 조회 중 입니다 ...', len(self.DATA['BL_LIST']))
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestLookUpDataFromUnipass',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '유니패스 조회 중 오류가 발생하였습니다.')

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
                            _bn_list=self.DATA['FILTER']['BN_LIST'],
                            _en_list=self.DATA['FILTER']['EN_LIST'],
                            _mn_list=self.DATA['FILTER']['MN_LIST'],
                            _exporter_shipper_name=self.DATA['FILTER']['EXPORTER_SHIPPER'],
                            _lookup_date_start=self.DATA['FILTER']['LOOKUP_DATE_START'],
                            _lookup_date_close=self.DATA['FILTER']['LOOKUP_DATE_CLOSE'],
                            _accept_date_start=self.DATA['FILTER']['ACCEPT_DATE_START'],
                            _accept_date_close=self.DATA['FILTER']['ACCEPT_DATE_CLOSE'],
                            _load_duty_date_start=self.DATA['FILTER']['LOAD_DUTY_DATE_START'],
                            _load_duty_date_close=self.DATA['FILTER']['LOAD_DUTY_DATE_CLOSE'],
                            _departure_date_start=self.DATA['FILTER']['DEPARTURE_DATE_START'],
                            _departure_date_close=self.DATA['FILTER']['DEPARTURE_DATE_CLOSE'],
                            _shipment_place=self.DATA['FILTER']['SHIPMENT_PLACE'],
                            _pre_shipping_flag=self.DATA['FILTER']['PRE_SHIPPING_FLAG'],
                            _manager_index=self.DATA['FILTER']['MANAGER_INDEX'],
                            _list_type=self.DATA['FILTER']['LIST_TYPE'],
                            _order_str=self.WIDGETS['TABLE'].get_head_sort()
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