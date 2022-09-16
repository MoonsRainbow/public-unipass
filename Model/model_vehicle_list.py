import sys
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetCategory(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelVehicleList_GetCategory'
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _category_list = []
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''SELECT manager_index, id FROM manager WHERE manager_index != 2;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                _query = '''SELECT vehicle_status FROM export_info_chassis_number GROUP BY vehicle_status;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                _query = '''SELECT flag_inspection FROM export_info_export_number GROUP BY flag_inspection;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                self.send(False, True, _category_list)
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()


class GetVehicleList(_ThreadModel):
    def __init__(
            self,
            _thread_queue: Queue,
            _en_list: list,
            _cn_list: list,
            _vehicle_status: str,
            _inspection_flag: str,
            _lookup_date_start: str,
            _lookup_date_close: str,
            _report_date_start: str,
            _report_date_close: str,
            _load_duty_date_start: str,
            _load_duty_date_close: str,
            _exporter_shipper_name: str,
            _manager_index: int,
            _list_type: int,
            _view_count: int,
            _page_no: int,
            _order_str: str,
            _app_queue: [Queue, None] = None
    ):
        self._code = 'ModelVehicleList_GetVehicleList'
        self._en_list = _en_list
        self._cn_list = _cn_list
        self._vehicle_status = _vehicle_status
        self._inspection_flag = _inspection_flag
        self._lookup_date_start = _lookup_date_start
        self._lookup_date_close = _lookup_date_close
        self._report_date_start = _report_date_start
        self._report_date_close = _report_date_close
        self._load_duty_date_start = _load_duty_date_start
        self._load_duty_date_close = _load_duty_date_close
        self._exporter_shipper_name = _exporter_shipper_name
        self._manager_index = _manager_index
        self._list_type = _list_type
        self._view_count = _view_count
        self._page_no = _page_no
        self._order_str = _order_str
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _response = []
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT COUNT(*), FLOOR(COUNT(*) / {}) + IF(MOD(COUNT(*), {}) > 0, 1, 0)
                FROM merged_ei_cn_en
                WHERE 1 = 1
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                {};
                '''.format(
                    self._view_count,
                    self._view_count,
                    'AND (KEY_INDEX) IN (SELECT MAX(KEY_INDEX) FROM merged_ei_cn_en GROUP BY CHASSIS_NUMBER, ORIGIN_EXPORT_NUMBER)' if self._list_type == 0 else '',
                    'AND (' + ' OR '.join(['ORIGIN_EXPORT_NUMBER LIKE \'%{}%\''.format(str(_).replace('-', '')) for _ in self._en_list]) + ')' if len(self._en_list) != 0 else '',
                    'AND (' + ' OR '.join(['CHASSIS_NUMBER LIKE \'%{}%\''.format(_) for _ in self._cn_list]) + ')' if len(self._cn_list) != 0 else '',
                    'AND VEHICLE_STATUS = \'{}\''.format(self._vehicle_status) if self._vehicle_status != 'All' else '',
                    'AND FLAG_INSPECTION = \'{}\''.format(self._inspection_flag) if self._inspection_flag != 'All' else '',
                    'AND (\'{}\' <= DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') AND DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') <= \'{}\')'.format(self._lookup_date_start, self._lookup_date_close),
                    'AND (\'{}\' <= REPORT_DATE AND REPORT_DATE <= \'{}\')'.format(self._report_date_start, self._report_date_close) if self._report_date_start != '' else '',
                    'AND (\'{}\' <= LOAD_DUTY_DEADLINE AND LOAD_DUTY_DEADLINE <= \'{}\')'.format(self._load_duty_date_start, self._load_duty_date_close) if self._load_duty_date_start != '' else '',
                    'AND EXPORT_SHIPPER LIKE \'%{}%\''.format(self._exporter_shipper_name) if self._exporter_shipper_name != '' else '',
                    'AND MANAGER_INDEX = {}'.format(self._manager_index) if self._manager_index != 0 else '',
                )
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _response.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                if len(_response) > 0:
                    _query = '''
                    SELECT KEY_INDEX,
                           FORMAT_EXPORT_NUMBER,
                           CHASSIS_NUMBER,
                           DATE_FORMAT(LOOKUP_DATE, '%Y-%m-%d'),
                           REPORT_DATE,
                           LOAD_DUTY_DEADLINE,
                           VEHICLE_STATUS,
                           TOTAL_AMOUNT,
                           TOTAL_WEIGHT,
                           PRE_AMOUNT,
                           PRE_WEIGHT,
                           REMAIN_AMOUNT,
                           REMAIN_WEIGHT,
                           FLAG_INSPECTION
                    FROM merged_ei_cn_en
                    WHERE 1 = 1
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    {}
                    '''.format(
                        'AND (KEY_INDEX) IN (SELECT MAX(KEY_INDEX) FROM merged_ei_cn_en GROUP BY CHASSIS_NUMBER, ORIGIN_EXPORT_NUMBER)' if self._list_type == 0 else '',
                        'AND (' + ' OR '.join(['ORIGIN_EXPORT_NUMBER LIKE \'%{}%\''.format(str(_).replace('-', '')) for _ in self._en_list]) + ')' if len(self._en_list) != 0 else '',
                        'AND (' + ' OR '.join(['CHASSIS_NUMBER LIKE \'%{}%\''.format(_) for _ in self._cn_list]) + ')' if len(self._cn_list) != 0 else '',
                        'AND VEHICLE_STATUS = \'{}\''.format(self._vehicle_status) if self._vehicle_status != 'All' else '',
                        'AND FLAG_INSPECTION = \'{}\''.format(self._inspection_flag) if self._inspection_flag != 'All' else '',
                        'AND (\'{}\' <= DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') AND DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') <= \'{}\')'.format(self._lookup_date_start, self._lookup_date_close),
                        'AND (\'{}\' <= REPORT_DATE AND REPORT_DATE <= \'{}\')'.format(self._report_date_start, self._report_date_close) if self._report_date_start != '' else '',
                        'AND (\'{}\' <= LOAD_DUTY_DEADLINE AND LOAD_DUTY_DEADLINE <= \'{}\')'.format(self._load_duty_date_start, self._load_duty_date_close) if self._load_duty_date_start != '' else '',
                        'AND EXPORT_SHIPPER LIKE \'%{}%\''.format(self._exporter_shipper_name) if self._exporter_shipper_name != '' else '',
                        'AND MANAGER_INDEX = {}'.format(self._manager_index) if self._manager_index != 0 else '',
                        self._order_str if self._order_str != '' else 'ORDER BY KEY_INDEX DESC',
                        'LIMIT {} OFFSET {};'.format(self._view_count, self._view_count * (self._page_no - 1))
                    )
                    _der = self._dbm.execute_query(_query)
                    if _der[0]:
                        _response.append(list(_der[1]).copy())
                        self.send(False, True, _response)
                    else:
                        self.send(False, False, _der[2])
            else:
                self.send(False, False, _dcr[1])
        except:
            self.send(False, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()
