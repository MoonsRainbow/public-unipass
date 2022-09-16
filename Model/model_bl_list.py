import sys
import time
import requests
from queue import Queue
from xml.etree import ElementTree
from Model import (_ThreadModel, _DataBaseManager)


class GetCategory(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelBlList_GetCategory'
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
                _query = '''SELECT shipment_place FROM export_info_bl_number GROUP BY shipment_place;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                _query = '''SELECT flag_pre_shipping FROM export_info_bl_number GROUP BY flag_pre_shipping;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                    return
                _query = '''SELECT env_value FROM admin_env WHERE env_type = 2;'''
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


class GetBlList(_ThreadModel):
    def __init__(
            self,
            _thread_queue: Queue,
            _bn_list: list,
            _en_list: list,
            _mn_list: list,
            _exporter_shipper_name: str,
            _lookup_date_start: str,
            _lookup_date_close: str,
            _accept_date_start: str,
            _accept_date_close: str,
            _load_duty_date_start: str,
            _load_duty_date_close: str,
            _departure_date_start: str,
            _departure_date_close: str,
            _shipment_place: str,
            _pre_shipping_flag: str,
            _manager_index: int,
            _list_type: int,
            _view_count: int,
            _page_no: int,
            _order_str: str,
            _app_queue: [Queue, None] = None
    ):
        self._code = 'ModelBlList_GetBlList'
        self._bn_list = _bn_list
        self._en_list = _en_list
        self._mn_list = _mn_list
        self._exporter_shipper_name = _exporter_shipper_name
        self._lookup_date_start = _lookup_date_start
        self._lookup_date_close = _lookup_date_close
        self._accept_date_start = _accept_date_start
        self._accept_date_close = _accept_date_close
        self._load_duty_date_start = _load_duty_date_start
        self._load_duty_date_close = _load_duty_date_close
        self._departure_date_start = _departure_date_start
        self._departure_date_close = _departure_date_close
        self._shipment_place = _shipment_place
        self._pre_shipping_flag = _pre_shipping_flag
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
                FROM merged_ei_bn
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
                {};
                '''.format(
                    self._view_count,
                    self._view_count,
                    'AND (KEY_INDEX) IN (SELECT MAX(KEY_INDEX) FROM merged_ei_bn GROUP BY BL_NUMBER, ORIGIN_EXPORT_NUMBER)' if self._list_type == 0 else '',
                    'AND (' + ' OR '.join(['BL_NUMBER LIKE \'%{}%\''.format(_) for _ in self._bn_list]) + ')' if len(self._bn_list) != 0 else '',
                    'AND (' + ' OR '.join(['ORIGIN_EXPORT_NUMBER LIKE \'%{}%\''.format(str(_).replace('-', '')) for _ in self._en_list]) + ')' if len(self._en_list) != 0 else '',
                    'AND (' + ' OR '.join(['MANAGEMENT_NUMBER LIKE \'%{}%\''.format(_) for _ in self._mn_list]) + ')' if len(self._mn_list) != 0 else '',
                    'AND EXPORT_SHIPPER LIKE \'%{}%\''.format(self._exporter_shipper_name) if self._exporter_shipper_name != '' else '',
                    'AND (\'{}\' <= DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') AND DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') <= \'{}\')'.format(self._lookup_date_start, self._lookup_date_close),
                    'AND (\'{}\' <= ACCEPT_DATE AND ACCEPT_DATE <= \'{}\')'.format(self._accept_date_start, self._accept_date_close) if self._accept_date_start != '' else '',
                    'AND (\'{}\' <= LOAD_DUTY_DEADLINE AND LOAD_DUTY_DEADLINE <= \'{}\')'.format(self._load_duty_date_start, self._load_duty_date_close) if self._load_duty_date_start != '' else '',
                    'AND (\'{}\' <= DEPARTURE_DATE AND DEPARTURE_DATE <= \'{}\')'.format(self._departure_date_start, self._departure_date_close) if self._departure_date_start != '' else '',
                    'AND SHIPMENT_PLACE LIKE \'%{}%\''.format(self._shipment_place) if self._shipment_place != 'All' else '',
                    'AND FLAG_PRE_SHIPPING LIKE \'%{}%\''.format(self._pre_shipping_flag) if self._pre_shipping_flag != 'All' else '',
                    'AND MANAGER_INDEX = {}'.format(self._manager_index) if self._manager_index != 0 else ''
                )
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _response.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
                if len(_response) > 0:
                    _query = '''
                    SELECT KEY_INDEX,
                           BL_NUMBER,
                           FORMAT_EXPORT_NUMBER,
                           MANAGEMENT_NUMBER,
                           LOOKUP_DATE,
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
                    {}
                    {};
                    '''.format(
                        'AND (KEY_INDEX) IN (SELECT MAX(KEY_INDEX) FROM merged_ei_bn GROUP BY BL_NUMBER, ORIGIN_EXPORT_NUMBER)' if self._list_type == 0 else '',
                        'AND (' + ' OR '.join(['BL_NUMBER LIKE \'%{}%\''.format(_) for _ in self._bn_list]) + ')' if len(self._bn_list) != 0 else '',
                        'AND (' + ' OR '.join(['ORIGIN_EXPORT_NUMBER LIKE \'%{}%\''.format(str(_).replace('-', '')) for _ in self._en_list]) + ')' if len(self._en_list) != 0 else '',
                        'AND (' + ' OR '.join(['MANAGEMENT_NUMBER LIKE \'%{}%\''.format(_) for _ in self._mn_list]) + ')' if len(self._mn_list) != 0 else '',
                        'AND EXPORT_SHIPPER LIKE \'%{}%\''.format(self._exporter_shipper_name) if self._exporter_shipper_name != '' else '',
                        'AND (\'{}\' <= DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') AND DATE_FORMAT(LOOKUP_DATE, \'%Y-%m-%d\') <= \'{}\')'.format(self._lookup_date_start, self._lookup_date_close),
                        'AND (\'{}\' <= ACCEPT_DATE AND ACCEPT_DATE <= \'{}\')'.format(self._accept_date_start, self._accept_date_close) if self._accept_date_start != '' else '',
                        'AND (\'{}\' <= LOAD_DUTY_DEADLINE AND LOAD_DUTY_DEADLINE <= \'{}\')'.format(self._load_duty_date_start, self._load_duty_date_close) if self._load_duty_date_start != '' else '',
                        'AND (\'{}\' <= DEPARTURE_DATE AND DEPARTURE_DATE <= \'{}\')'.format(self._departure_date_start, self._departure_date_close) if self._departure_date_start != '' else '',
                        'AND SHIPMENT_PLACE LIKE \'%{}%\''.format(self._shipment_place) if self._shipment_place != 'All' else '',
                        'AND FLAG_PRE_SHIPPING LIKE \'%{}%\''.format(self._pre_shipping_flag) if self._pre_shipping_flag != 'All' else '',
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


class LookUpDataFromUnipass(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _bl_list: list, _app_queue: [Queue, None] = None):
        self._code, self._manager_index, self._bl_list, self._lookup_index = 'ModelBlList_LookUpDataFromUnipass', _manager_index, _bl_list, 0
        self._unipass_bn_key = ''
        self._unipass_bn_url = ''
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''INSERT INTO lookup_bl(manager_index) VALUES({})'''.format(self._manager_index)
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self._lookup_index = _der[1]
                    for _origin_bl_info in self._bl_list:
                        _is_critical = False
                        _is_warning = False
                        _response_info = {}
                        _remark = {
                            'OBN': str(_origin_bl_info[0]),
                            'URL': None,
                            'DESC': None
                        }
                        _lookup_data_list = []
                        _bn_lookup_data = {
                            'ERN': int(_origin_bl_info[1]),
                            'OBN': str(_origin_bl_info[0])
                        }
                        try:
                            _bn_response = None
                            _remark['URL'] = self._unipass_bn_url.format(
                                self._unipass_bn_key, _bn_lookup_data['OBN']
                            )
                            for _ in range(11):
                                try:
                                    _bn_response = requests.get(_remark['URL'])
                                    break
                                except requests.exceptions.ConnectionError:
                                    if _ == 10:
                                        _is_critical = True
                                        _remark['DESC'] = '<requests.exceptions.ConnectionError> occurred while processing the Unipass API.'
                                        _response_info = {
                                            'ECD': self._code + '_Run',
                                            'CLS': sys.exc_info()[0],
                                            'DES': sys.exc_info()[1],
                                            'LNO': sys.exc_info()[2].tb_lineno,
                                            'DTA': [_remark]
                                        }
                                        _bn_response = None
                                        break
                                    else:
                                        try:
                                            time.sleep(2)
                                            _bn_response = requests.get(_remark['URL'])
                                            break
                                        except requests.exceptions.ConnectionError:
                                            _bn_response = None
                                            continue
                            if not _is_critical:
                                if _bn_response.status_code == 200:
                                    _bn_xml = ElementTree.fromstring(_bn_response.content.decode('utf-8'))
                                    _bn_count = int(_bn_xml.find('tCnt').text)
                                    if _bn_count == -1:
                                        _is_warning = True
                                        _response_info = {
                                            'ECD': self._code + '_Run',
                                            'CLS': 'Abnormal Result',
                                            'DES': _bn_xml.find('ntceInfo').text,
                                            'LNO': 0,
                                            'DTA': [_remark]
                                        }
                                    elif _bn_count == 0:
                                        _bn_lookup_data['BN'] = _bn_lookup_data.get('OBN')
                                        _bn_lookup_data['EN'] = None
                                        _lookup_data_list.append(_bn_lookup_data.copy())
                                    else:
                                        for _bn_data in _bn_xml.findall('expDclrNoPrExpFfmnBrkdBlNoQryRsltVo'):
                                            _bn_lookup_data['BN'] = _bn_lookup_data.get('OBN')
                                            _bn_lookup_data['AD'] = str(_bn_data.find('acptDt').text)[:4] + '-' + str(_bn_data.find('acptDt').text)[4:6] + '-' + str(_bn_data.find('acptDt').text)[6:]
                                            _bn_lookup_data['ES'] = _bn_data.findall('exppnConm')[0].text
                                            _bn_lookup_data['LDD'] = str(_bn_data.find('loadDtyTmlm').text)[:4] + '-' + str(_bn_data.find('loadDtyTmlm').text)[4:6] + '-' + str(_bn_data.find('loadDtyTmlm').text)[6:]
                                            _bn_lookup_data['SP'] = _bn_data.findall('shpmAirptPortNm')[0].text
                                            _bn_lookup_data['DD'] = str(_bn_data.find('tkofDt').text)[:4] + '-' + str(_bn_data.find('tkofDt').text)[4:6] + '-' + str(_bn_data.find('tkofDt').text)[6:]
                                            _bn_lookup_data['MN'] = _bn_data.findall('mrn')[0].text
                                            _bn_lookup_data['PA'] = int(_bn_data.findall('shpmPckGcnt')[0].text)
                                            _bn_lookup_data['PW'] = int(float(_bn_data.find('shpmWght').text))
                                            _bn_lookup_data['FPS'] = _bn_data.findall('shpmCmplYn')[0].text
                                            _bn_lookup_data['DC'] = int(_bn_data.findall('dvdeWdrw')[0].text)
                                            _bn_lookup_data['EN'] = _bn_data.findall('expDclrNo')[0].text
                                            _bn_lookup_data['TA'] = int(_bn_data.findall('csclPckGcnt')[0].text)
                                            _bn_lookup_data['TW'] = int(float(_bn_data.find('csclWght').text))
                                            _lookup_data_list.append(_bn_lookup_data.copy())
                                else:
                                    _is_critical = True
                                    _remark['DESC'] = 'The Unipass API sent a {} code.'.format(_bn_response.status_code)
                                    _response_info = {
                                        'ECD': self._code + '_Run',
                                        'CLS': 'Non-200 Code',
                                        'DES': _bn_response.status_code,
                                        'LNO': 0,
                                        'DTA': [_remark]
                                    }
                                    break
                            else:
                                break
                        except:
                            _is_critical = True
                            _remark['DESC'] = 'An error occurred while processing the Unipass API. Debugging required.'
                            _response_info = {
                                'ECD': self._code + '_Run',
                                'CLS': sys.exc_info()[0],
                                'DES': sys.exc_info()[1],
                                'LNO': sys.exc_info()[2].tb_lineno,
                                'DTA': [_remark]
                            }
                        finally:
                            if _is_critical:
                                self.send(True, False, _response_info)
                                break
                            else:
                                if _is_warning:
                                    self.send(True, False, _response_info)
                                else:
                                    for _lookup_data in _lookup_data_list:
                                        _remark['OBN'] = _lookup_data['OBN']
                                        _remark['LBN'] = _lookup_data['BN']
                                        _remark['LEN'] = _lookup_data['EN']
                                        _remark.__setitem__('DESC', None)
                                        _remark.__setitem__('URL', None)
                                        _query = 'INSERT INTO dict_bl_number(bl_number) VALUES(\'{}\')'.format(_lookup_data['BN'])
                                        _der = self._dbm.execute_query(_query)
                                        if _der[0]:
                                            _d_bn_index = _der[1]
                                        else:
                                            if 'Duplicate entry' in str(_der[2]['DES']):
                                                _query2 = 'SELECT d_bl_index FROM dict_bl_number WHERE bl_number = \'{}\''.format(_lookup_data['BN'])
                                                _der2 = self._dbm.execute_query(_query2)
                                                if _der2[0]:
                                                    _d_bn_index = _der2[1][0][0]
                                                else:
                                                    _der2[2]['DTA'].append(_remark.copy())
                                                    self.send(True, False, _der2[2])
                                                    break
                                            else:
                                                _der[2]['DTA'].append(_remark.copy())
                                                self.send(True, False, _der[2])
                                                break
                                        if _lookup_data['EN'] is not None:
                                            _query = 'INSERT INTO dict_export_number(export_number) VALUES(\'{}\')'.format(_lookup_data['EN'])
                                            _der = self._dbm.execute_query(_query)
                                            if _der[0]:
                                                _d_en_index = _der[1]
                                            else:
                                                if 'Duplicate entry' in str(_der[2]['DES']):
                                                    _query2 = 'SELECT d_export_index FROM dict_export_number WHERE export_number = \'{}\''.format(_lookup_data['EN'])
                                                    _der2 = self._dbm.execute_query(_query2)
                                                    if _der2[0]:
                                                        _d_en_index = _der2[1][0][0]
                                                    else:
                                                        _der2[2]['DTA'].append(_remark.copy())
                                                        self.send(True, False, _der2[2])
                                                        break
                                                else:
                                                    _der[2]['DTA'].append(_remark.copy())
                                                    self.send(True, False, _der[2])
                                                    break
                                            _query = '''
                                            INSERT INTO export_info_bl_number(
                                                d_export_index,
                                                management_number,
                                                accept_date, load_duty_deadline,
                                                export_shipper,
                                                total_amount,
                                                total_weight,
                                                shipment_place,
                                                departure_date,
                                                pre_shipping_amount,
                                                pre_shipping_weight,
                                                flag_pre_shipping,
                                                division_collect)
                                            VALUES ({},
                                                    '{}',
                                                    '{}',
                                                    '{}',
                                                    '{}',
                                                    {},
                                                    {},
                                                    '{}',
                                                    '{}',
                                                    {},
                                                    {},
                                                    '{}',
                                                    {})'''.format(
                                                _d_en_index,
                                                _lookup_data['MN'],
                                                _lookup_data['AD'],
                                                _lookup_data['LDD'],
                                                _lookup_data['ES'],
                                                _lookup_data['TA'],
                                                _lookup_data['TW'],
                                                _lookup_data['SP'],
                                                _lookup_data['DD'],
                                                _lookup_data['PA'],
                                                _lookup_data['PW'],
                                                _lookup_data['FPS'],
                                                _lookup_data['DC']
                                            )
                                            _der = self._dbm.execute_query(_query)
                                            if _der[0]:
                                                _ei_bl_index = _der[1]
                                            else:
                                                _der[2]['DTA'].append(_remark.copy())
                                                self.send(True, False, _der[2])
                                                break
                                        else:
                                            _d_en_index = 0
                                            _ei_bl_index = 0
                                        _query = '''
                                        INSERT INTO lookup_bl_result(l_bl_index, d_bl_index, ei_bl_index, excel_row)
                                        VALUES({}, {}, {}, {})'''.format(
                                            self._lookup_index, _d_bn_index, _ei_bl_index, _lookup_data['ERN']
                                        )
                                        _der = self._dbm.execute_query(_query)
                                        if not _der[0]:
                                            _der[2]['DTA'].append(_remark.copy())
                                            self.send(True, False, _der[2])
                                            break
                                    else:
                                        self.send(True, True, None)
                            if self.recv() is not None:
                                break
                else:
                    self.send(True, False, _der[2])
            else:
                self.send(True, False, _dcr[1])
        except:
            self.send(True, False, {
                'ECD': self._code + '_Run',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            })
        finally:
            self._dbm.disconnecting()
            self.send(False, False, None)

class GetBlDetailList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _key_index: int, _app_queue: [Queue, None] = None):
        self._code, self._key_index = 'ModelBlList_GetBlDetailList', _key_index
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
