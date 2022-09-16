import sys
from queue import Queue
from Model import (_ThreadModel, _DataBaseManager)


class GetCategory(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelLookUpBlHistory_GetCategory'
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
        self._code, self._lookup_date_start, self._lookup_date_close, self._manager_index = 'ModelLookUpBlHistory_GetLookUpList', _lookup_date_start, _lookup_date_close, _manager_index
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT l_bl_index,
                       m.id,
                       DATE_FORMAT(lookup_date, '%Y-%m-%d %h:%i:%s'),
                       (SELECT COUNT(*) FROM lookup_bl_result lbr WHERE lbr.l_bl_index = lb.l_bl_index)
                FROM lookup_bl lb
                         LEFT JOIN manager m on lb.manager_index = m.manager_index
                WHERE 
                {}
                {}
                ORDER BY l_bl_index DESC;
                '''.format(
                    "(lookup_date BETWEEN '{0}' AND '{1}' OR lookup_date LIKE '{0}%' OR lookup_date LIKE '{1}%')".format(self._lookup_date_start, self._lookup_date_close) if self._lookup_date_start != '' else '',
                    'AND lb.manager_index = {}'.format(self._manager_index) if self._manager_index != 0 else ''
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


class GetBlList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _lookup_index: int, _order_str: str, _app_queue: [Queue, None] = None):
        self._code, self._lookup_index, self._order_str = 'ModelLookUpBlHistory_GetBlList', _lookup_index, _order_str
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''
                SELECT KEY_INDEX,
                       EXCEL_ROW_NO,
                       BL_NUMBER,
                       FORMAT_EXPORT_NUMBER,
                       MANAGEMENT_NUMBER,
                       ACCEPT_DATE,
                       LOAD_DUTY_DEADLINE,
                       EXPORT_SHIPPER,
                       TOTAL_AMOUNT,
                       TOTAL_WEIGHT,
                       SHIPMENT_PLACE,
                       DEPARTURE_DATE,
                       PRE_AMOUNT,
                       PRE_WEIGHT,
                       DIVISION_COLLECT,
                       FLAG_PRE_SHIPPING
                FROM merged_ei_bn
                WHERE LOOKUP_BL_INDEX = {}
                {};
                '''.format(
                    self._lookup_index,
                    self._order_str if self._order_str is not None else 'ORDER BY KEY_INDEX DESC'
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


class GetBlDetailList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _key_index: int, _app_queue: [Queue, None] = None):
        self._code, self._key_index = 'ModelLookUpBlHistory_GetBlDetailList', _key_index
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = f'''
                SELECT lb.LOOKUP_DATE,
                       BL_NUMBER,
                       EXPORT_SHIPPER,
                       ACCEPT_DATE,
                       TOTAL_AMOUNT,
                       SHIPMENT_PLACE,
                       PRE_AMOUNT,
                       DIVISION_COLLECT,
                       FORMAT_EXPORT_NUMBER,
                       LOAD_DUTY_DEADLINE,
                       TOTAL_WEIGHT,
                       DEPARTURE_DATE,
                       PRE_WEIGHT,
                       FLAG_PRE_SHIPPING,
                       MANAGEMENT_NUMBER
                FROM merged_ei_bn JOIN lookup_bl lb ON merged_ei_bn.LOOKUP_BL_INDEX = lb.l_bl_index
                WHERE D_BL_INDEX = (SELECT D_BL_INDEX FROM merged_ei_bn WHERE KEY_INDEX = {self._key_index})
                  AND LOOKUP_BL_INDEX = (SELECT LOOKUP_BL_INDEX FROM merged_ei_bn WHERE KEY_INDEX = {self._key_index})
                ORDER BY KEY_INDEX;'''
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