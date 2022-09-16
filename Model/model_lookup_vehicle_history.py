import sys
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetCategory(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelLookUpVehicleHistory_GetCategory'
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT manager_index, id
                FROM manager
                WHERE manager_index != 2;
                '''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True, _der[1])
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


class GetLookUpList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _lookup_date_start: str, _lookup_date_close: str, _manager_index: int, _app_queue: [Queue, None] = None):
        self._code, self._lookup_date_start, self._lookup_date_close, self._manager_index = 'ModelLookUpVehicleHistory_GetLookUpList', _lookup_date_start, _lookup_date_close, _manager_index
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT l_chassis_index,
                       m.id,
                       DATE_FORMAT(lookup_date, '%Y-%m-%d %h:%i:%s'),
                       (SELECT COUNT(*) FROM lookup_chassis_result lcr WHERE lcr.l_chassis_index = lc.l_chassis_index)
                FROM lookup_chassis lc
                         LEFT JOIN manager m on lc.manager_index = m.manager_index
                WHERE 
                {}
                {}
                ORDER BY l_chassis_index DESC;
                '''.format(
                    "(lookup_date BETWEEN '{0}' AND '{1}' OR lookup_date LIKE '{0}%' OR lookup_date LIKE '{1}%')".format(self._lookup_date_start, self._lookup_date_close) if self._lookup_date_start != '' else '',
                    'AND lc.manager_index = {}'.format(self._manager_index) if self._manager_index != 0 else ''
                )
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True, _der[1])
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


class GetVehicleList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _lookup_index: int, _order_str: str, _app_queue: [Queue, None] = None):
        self._code, self._manager_index, self._lookup_index, self._order_str = 'ModelLookUpVehicleHistory_GetVehicleList', _manager_index, _lookup_index, _order_str
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT KEY_INDEX,
                       GET_COMPARE_LOOKUP_CHASSIS_RESULT(KEY_INDEX, {}) AS RESULT,
                       EXCEL_ROW_NO,
                       FORMAT_EXPORT_NUMBER,
                       CHASSIS_NUMBER,
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
                WHERE LOOKUP_CHASSIS_INDEX = {}
                {};
                '''.format(
                    self._manager_index,
                    self._lookup_index,
                    self._order_str if self._order_str is not None else 'ORDER BY EXCEL_ROW_NO DESC'
                )
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self.send(False, True, _der[1])
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
