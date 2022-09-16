import os
import queue
import sys
import xlrd
import tkinter
import datetime
from tkinter import messagebox, filedialog
from Model.model_common import GetLookUpDataDetail, DownloadVehicleDataToExcel
from View import Body, LabelEntryButton, LabelText, LabelCombo, Table
from Model.model_vehicle_lookup import GetCategory, LookUpDataFromUnipass, GetLookUpDataList


class VehicleLookUpBody(Body):
    def __init__(self, _app):
        self.APP, self.WIDGETS, self.DATA = _app, {}, {}
        super().__init__(_app)
        self._code = 'ViewVehicleLookUp_VehicleLookUpBody'

    def initial(self):
        try:
            self.destroy_widget()
            self.DATA['CATEGORY_MAN_ID'] = {'All': 0}
            self.DATA['CATEGORY_COLUMN_TITLE'] = []
            self.DATA['CHASSIS_LIST'] = []
            self.DATA['CURRENT_I_INDEX'] = 0
            self.WIDGETS.clear()
            self.request_get_category()
            self.WIDGETS['LEB_CHASSIS'] = LabelEntryButton(self, '차대번호', '등록', self.event_button_click_chassis_add)
            self.WIDGETS['B_EXCEL'] = tkinter.Button(self, text='엑셀 선택', overrelief=tkinter.SOLID, command=self.event_button_click_excel_choice, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['TABLE'] = Table(self.APP, self)
            self.WIDGETS['LT_TOTAL'] = LabelText(self, '전체 개수', self.WIDGETS['TABLE'].COUNT['TOT'])
            self.WIDGETS['LT_NEW'] = LabelText(self, '신규 개수', self.WIDGETS['TABLE'].COUNT['NEW'])
            self.WIDGETS['LT_SAME'] = LabelText(self, '동일 개수', self.WIDGETS['TABLE'].COUNT['SAM'])
            self.WIDGETS['LT_CHANGE'] = LabelText(self, '변경 개수', self.WIDGETS['TABLE'].COUNT['CHG'])
            self.WIDGETS['LC_ID'] = LabelCombo(self, '조회 아이디', self.DATA['CATEGORY_MAN_ID'])
            self.WIDGETS['B_COPY'] = tkinter.Button(self, textvariable=self.WIDGETS['TABLE'].CELL_SELECTION_COUNT, overrelief=tkinter.SOLID, command=self.event_button_click_selection_copy, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_DOWNLOAD'] = tkinter.Button(self, text='엑셀\n다운로드', overrelief=tkinter.SOLID, command=self.event_button_click_excel_download, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_RESET'] = tkinter.Button(self, text='목록\n초기화', overrelief=tkinter.SOLID, command=self.event_button_click_reset, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['B_LOOKUP'] = tkinter.Button(self, text='UNIPASS 조회', overrelief=tkinter.SOLID, command=self.event_button_click_data_lookup_at_unipass, repeatdelay=1000, repeatinterval=100)
            self.WIDGETS['LEB_CHASSIS'].place(relwidth=0.871, height=56)
            self.WIDGETS['B_EXCEL'].place(relx=0.88, relwidth=0.12, height=56)
            self.WIDGETS['LT_TOTAL'].place(height=56, relwidth=0.07, y=66)
            self.WIDGETS['LT_NEW'].place(height=56, relwidth=0.07, y=66, relx=0.079)
            self.WIDGETS['LT_SAME'].place(height=56, relwidth=0.07, y=66, relx=0.158)
            self.WIDGETS['LT_CHANGE'].place(height=56, relwidth=0.07, y=66, relx=0.237)
            self.WIDGETS['LC_ID'].place(height=56, relwidth=0.12, y=66, relx=0.316)
            self.WIDGETS['B_COPY'].place(height=56, relwidth=0.06, y=66, relx=0.673)
            self.WIDGETS['B_DOWNLOAD'].place(height=56, relwidth=0.06, y=66, relx=0.742)
            self.WIDGETS['B_RESET'].place(height=56, relwidth=0.06, y=66, relx=0.811)
            self.WIDGETS['B_LOOKUP'].place(height=56, relwidth=0.12, y=66, relx=0.88)
            self.WIDGETS['TABLE'].set_location_y(152)
            self.WIDGETS['TABLE'].LOCATION.place(y=132, relwidth=1, height=self.APP.get_current_geometry()[1] * 0.94 - 152)
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
            self.APP.bind('<Control-c>', self.event_button_click_selection_copy)
            self.APP.bind('<Return>', self.event_button_click_chassis_add)
            self.DATA['CHASSIS_LIST'] = []
            self.DATA['CURRENT_I_INDEX'] = 0
            self.WIDGETS['LC_ID'].set_combo_state(tkinter.DISABLED)
            self.WIDGETS['LC_ID'].set_combo_selected_command(self.event_combo_select_id)
            self.WIDGETS['B_DOWNLOAD'].configure(state=tkinter.DISABLED)
            self.WIDGETS['B_LOOKUP'].configure(state=tkinter.DISABLED)
            self.WIDGETS['TABLE'].set_table_resizing()
            self.WIDGETS['TABLE'].set_table_paging()
            self.WIDGETS['TABLE'].set_table_celling(True)
            self.WIDGETS['TABLE'].set_head_sort_command()
            self.WIDGETS['TABLE'].set_cell_double_click_command()
            self.WIDGETS['TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.COMMON_HEAD_CATEGORY.items() if 'VLB' in _value['TABLE']},
                _is_change=True
            )
            self.WIDGETS['TABLE'].cell_clear()
            self.WIDGETS['TABLE'].clear_count()
            self.WIDGETS['TABLE'].delete_row()
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

    def view_lookup_data(self):
        try:
            self.APP.bind('<Control-c>', self.event_button_click_selection_copy)
            self.APP.bind('<Return>', self.event_button_click_chassis_add)
            self.DATA['CHASSIS_LIST'] = []
            self.WIDGETS['LC_ID'].set_combo_state()
            self.WIDGETS['B_DOWNLOAD'].configure(state=tkinter.NORMAL)
            self.WIDGETS['B_LOOKUP'].configure(state=tkinter.DISABLED)
            self.WIDGETS['TABLE'].set_head_sort_command(_sort_command=self.event_table_head_click)
            self.WIDGETS['TABLE'].set_cell_double_click_command(_double_click_command=self.event_table_cell_double_click)
            self.WIDGETS['TABLE'].head_set(
                _selected_head={_key: _value for _key, _value in self.APP.VEHICLE_HEAD_CATEGORY.items() if 'VLR' in _value['TABLE']},
                _is_change=True
            )
            self.WIDGETS['TABLE'].head_sort_clear()
            self.WIDGETS['TABLE'].cell_clear()
            self.WIDGETS['TABLE'].clear_count()
            self.WIDGETS['TABLE'].delete_row()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_ViewLookUpData',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '화면 초기화 중 오류가 발생하였습니다.')

    def event_button_click_chassis_add(self, _event=None):
        try:
            if len(self.WIDGETS['LEB_CHASSIS'].get_value()) != 0:
                if len(self.DATA['CHASSIS_LIST']) == 0:
                    self.view_reset()
                _inserted_cn_list = str(self.WIDGETS['LEB_CHASSIS'].get_value()).split(' ')
                _cn_list = []
                for _cn in _inserted_cn_list:
                    if '\t\n' in _cn:
                        _cn_list += _cn.split('\t\n')
                    elif '\t' in _cn:
                        _cn_list += _cn.split('\t')
                    elif '\n' in _cn:
                        _cn_list += _cn.split('\n')
                    else:
                        _cn_list.append(_cn)
                for _cn in _cn_list:
                    _cn = _cn.strip().upper()
                    if _cn in [chassis_info['OCN'] for chassis_info in self.DATA['CHASSIS_LIST']]:
                        self.WIDGETS['TABLE'].insert_row((('-', _cn, '중복된 차대번호 입니다.'),))
                    elif len(_cn) < 1:
                        pass
                    elif len(_cn) <= 7:
                        self.WIDGETS['TABLE'].insert_row((('-', _cn, '8자리 미만의 차대번호는 조회 시 제외됩니다.'),))
                    elif len(_cn) > 17:
                        self.WIDGETS['TABLE'].insert_row((('-', _cn, '17자리 초과의 차대번호는 조회 시 제외됩니다.'),))
                    else:
                        self.WIDGETS['TABLE'].set_count_total(self.WIDGETS['TABLE'].get_count_total() + 1)
                        self.WIDGETS['TABLE'].insert_row(((self.WIDGETS['TABLE'].get_count_total(), _cn, '-'),))
                        self.DATA['CHASSIS_LIST'].append({
                            'OCN': _cn,
                            'ERN': 0
                        })
            self.WIDGETS['LEB_CHASSIS'].set_value()
            if len(self.DATA['CHASSIS_LIST']) > 0:
                self.WIDGETS['B_LOOKUP'].configure(state=tkinter.NORMAL)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickChassisAdd',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '차대번호 등록 중 오류가 발생하였습니다.')

    def event_button_click_excel_choice(self):
        try:
            _excel_path = self.request_choice_excel()
            if _excel_path != '':
                if len(self.DATA['CHASSIS_LIST']) == 0:
                    self.view_reset()
                _cer = self.request_choice_sheet(_excel_path)
                if _cer[0]:
                    if _cer[1] is not None:
                        _rsr = self.request_read_sheet(_excel_path, _cer[1])
                        if _rsr[0]:
                            for _row in _rsr[1]:
                                if _row[0] != '-':
                                    self.WIDGETS['TABLE'].set_count_total(self.WIDGETS['TABLE'].get_count_total() + 1)
                                    _row[0] = self.WIDGETS['TABLE'].get_count_total()
                                    self.DATA['CHASSIS_LIST'].append({
                                        'OCN': _row[1],
                                        'ERN': _row.pop(3)
                                    })
                                self.WIDGETS['TABLE'].insert_row((_row,))
                        else:
                            if _rsr[1] is None:
                                messagebox.showwarning('UNIPASS ::', '\'{}\' 시트에서\n차대번호를 찾을 수 없습니다.\n확인 후 다시 시도해주세요.'.format(_cer[1]))
                            else:
                                self.APP.request_error_report_db(_error_info=_rsr[1])
                                messagebox.showerror('UNIPASS :: ', '엑셀 시트 읽기 중 오류가 발생하였습니다.')
                else:
                    self.APP.request_error_report_db(_error_info=_cer[1])
                    messagebox.showerror('UNIPASS :: ', '엑셀 시트 선택 중 오류가 발생하였습니다.')
                if len(self.DATA['CHASSIS_LIST']) > 0:
                    self.WIDGETS['B_LOOKUP'].configure(state=tkinter.NORMAL)
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
        self.view_reset()

    def event_button_click_data_lookup_at_unipass(self):
        try:
            self.APP.unbind('<Control-c>')
            self.APP.unbind('<Return>')
            self.request_lookup_data_from_unipass()
            if self.DATA['CURRENT_I_INDEX'] != 0:
                self.view_lookup_data()
                self.request_get_lookup_data_list()
            else:
                messagebox.showerror('UNIPASS :: ', '생성된 조회 번호가 없습니다.')
            self.APP.bind('<Control-c>', self.event_button_click_selection_copy)
            self.APP.bind('<Return>', self.event_button_click_chassis_add)
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventButtonClickDataLookUpAtUnipass',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '유니패스 조회 중 오류가 발생하였습니다.')

    def event_combo_select_id(self, _event=None):
        try:
            self.WIDGETS['TABLE'].head_sort_clear()
            self.WIDGETS['TABLE'].cell_clear()
            self.WIDGETS['TABLE'].clear_count()
            self.WIDGETS['TABLE'].delete_row()
            self.request_get_lookup_data_list()
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_EventComboSelectId',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '조회 아이디 변경 중 오류가 발생하였습니다.')

    def event_table_head_click(self, _head_key: str = ''):
        try:
            self.WIDGETS['TABLE'].head_sort(_head_key)
            self.WIDGETS['TABLE'].cell_clear()
            self.WIDGETS['TABLE'].clear_count()
            self.WIDGETS['TABLE'].delete_row()
            self.request_get_lookup_data_list()
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
                        for _row in _response['DATA'][1]:
                            self.DATA['CATEGORY_COLUMN_TITLE'].append(_row[0])
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '카테고리를 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = { 'THREAD': GetCategory(_thread_queue=self.APP.THREAD_QUEUE) }
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
                _cn_list = []
                for _row in range(_s_row + 1, _sheet.nrows):
                    _cn = str(_sheet.cell_value(_row, _s_col)).strip().upper()
                    if _cn in [chassis_info['OCN'] for chassis_info in self.DATA['CHASSIS_LIST']]:
                        _cn_list.append(['-', _cn, '중복된 차대번호 입니다.'])
                    elif _cn in [ _[1] for _ in _cn_list ]:
                        _cn_list.append(['-', _cn, '중복된 차대번호 입니다.'])
                    elif len(_cn) < 1:
                        pass
                    elif len(_cn) <= 7:
                        _cn_list.append(['-', _cn, '8자리 미만의 차대번호는 조회 시 제외됩니다.'])
                    elif len(_cn) > 17:
                        _cn_list.append(['-', _cn, '17자리 초과의 차대번호는 조회 시 제외됩니다.'])
                    else:
                        _cn_list.append([0, _cn, '-', _row + 1])
                return True, _cn_list
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
                            _inquiry_index=self.DATA['CURRENT_I_INDEX']
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

    def request_lookup_data_from_unipass(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_lookup_data_from_unipass, True)
                else:
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['RUN']:
                        if _response['SUCCESS']:
                            if _response['DATA'] is not None:
                                self.DATA['CURRENT_I_INDEX'] = _response['DATA']
                            else:
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
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': LookUpDataFromUnipass(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _app_queue=self.APP.APP_QUEUE,
                            _manager_index=int(self.APP.getvar('MI')),
                            _chassis_list=self.DATA['CHASSIS_LIST']
                        ),
                        'ERROR_LIST': []
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_lookup_data_from_unipass, True)
                    self.APP.POPUP.lookup_data('유니패스 조회 중 입니다 ...', self.WIDGETS['TABLE'].COUNT['TOT'])
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

    def request_get_lookup_data_list(self, _is_run: bool = False):
        try:
            if _is_run:
                if self.APP.THREAD_QUEUE.empty():
                    self.APP.after(100, self.request_get_lookup_data_list, True)
                else:
                    if self.APP.CURRENT_THREAD_INFO['THREAD'].isAlive():
                        self.APP.CURRENT_THREAD_INFO['THREAD'].join()
                    self.APP.CURRENT_THREAD_INFO.clear()
                    self.APP.POPUP.destroy()
                    _response = self.APP.THREAD_QUEUE.get()
                    if _response['SUCCESS']:
                        self.WIDGETS['TABLE'].delete_row()
                        self.WIDGETS['TABLE'].set_count_total(len(_response['DATA']))
                        self.WIDGETS['TABLE'].insert_row(_response['DATA'], is_tag=True)
                    else:
                        self.APP.request_error_report_db(_error_info=_response['DATA'])
                        messagebox.showerror('UNIPASS :: ', '차량 목록을 불러오는 중 오류가 발생하였습니다.')
            else:
                if len(self.APP.CURRENT_THREAD_INFO) == 0:
                    self.APP.clear_queues()
                    self.APP.CURRENT_THREAD_INFO = {
                        'THREAD': GetLookUpDataList(
                            _thread_queue=self.APP.THREAD_QUEUE,
                            _manager_index=self.WIDGETS['LC_ID'].get_value_index(),
                            _inquiry_index=self.DATA['CURRENT_I_INDEX'],
                            _order_str=self.WIDGETS['TABLE'].get_head_sort()
                        )
                    }
                    self.APP.CURRENT_THREAD_INFO['THREAD'].setDaemon(True)
                    self.APP.CURRENT_THREAD_INFO['THREAD'].start()
                    self.APP.after(100, self.request_get_lookup_data_list, True)
                    self.APP.POPUP.loading('차량 목록을 불러오는 중 입니다 ...')
        except:
            self.APP.request_error_report_db(
                _error_info={
                    'ECD': self._code + '_RequestGetLookUpDataList',
                    'CLS': sys.exc_info()[0],
                    'DES': sys.exc_info()[1],
                    'LNO': sys.exc_info()[2].tb_lineno,
                    'DTA': []
                }
            )
            messagebox.showerror('UNIPASS :: ', '차량 목록을 불러오는 중 오류가 발생하였습니다.')

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
