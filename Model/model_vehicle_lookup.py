import sys
import time
import requests
import datetime
from queue import Queue
from xml.etree import ElementTree
from Model import (_ThreadModel, _DataBaseManager)


class GetCategory(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _app_queue: [Queue, None] = None):
        self._code = 'ModelVehicleLookup_GetCategory'
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
                _query = '''SELECT env_value FROM admin_env WHERE env_type = 1;'''
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    _category_list.append(list(_der[1]).copy())
                else:
                    self.send(False, False, _der[2])
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


class LookUpDataFromUnipass(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _chassis_list: list, _app_queue: Queue, _max_year: int = 5):
        self._code, self._manager_index, self._chassis_list, self._max_year, self._inquiry_index = 'ModelVehicleLookup_LookUpDataFromUnipass', _manager_index, _chassis_list, _max_year, 0
        self._unipass_cn_key = ''
        self._unipass_en_key = ''
        self._unipass_cn_url = r''
        self._unipass_en_url = r''
        super().__init__(_thread_queue, _app_queue)
        self._dbm = _DataBaseManager(self._code)

    def run(self):
        try:
            _dcr = self._dbm.connecting()
            if _dcr[0]:
                _query = '''INSERT INTO lookup_chassis(manager_index) VALUES({})'''.format(self._manager_index)
                _der = self._dbm.execute_query(_query)
                if _der[0]:
                    self._inquiry_index = _der[1]
                    self.send(True, True, self._inquiry_index)
                    for _origin_chassis_info in self._chassis_list:
                        _is_critical = False
                        _is_warning = False
                        _response_info = {}
                        _remark = {
                            'OCN': str(_origin_chassis_info['OCN']),
                            'URL': None,
                            'DESC': None
                        }
                        _lookup_data_list = []
                        _cn_lookup_data = {
                            'ERN': int(_origin_chassis_info['ERN']),
                            'OCN': str(_origin_chassis_info['OCN']),
                            'EN': None,
                            'RD': None,
                            'VS': None
                        }
                        try:
                            _cn_response = None
                            for _year in range(1, self._max_year + 1):
                                cn_sd = '{}{}{}'.format(
                                    int(datetime.datetime.now().year) - _year,
                                    str(int(datetime.datetime.now().month)).zfill(2),
                                    str(int(datetime.datetime.now().day) + 1).zfill(2)
                                )
                                cn_ed = '{}{}{}'.format(
                                    int(datetime.datetime.now().year) - (_year - 1),
                                    str(int(datetime.datetime.now().month)).zfill(2),
                                    str(int(datetime.datetime.now().day)).zfill(2)
                                )
                                _remark['URL'] = self._unipass_cn_url.format(
                                    self._unipass_cn_key, cn_sd, cn_ed, _cn_lookup_data['OCN']
                                )
                                for _ in range(11):
                                    try:
                                        _cn_response = requests.get(_remark['URL'])
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
                                            _cn_response = None
                                            break
                                        else:
                                            try:
                                                time.sleep(2)
                                                _cn_response = requests.get(_remark['URL'])
                                                break
                                            except requests.exceptions.ConnectionError:
                                                _cn_response = None
                                                continue
                                if not _is_critical:
                                    if _cn_response.status_code == 200:
                                        _cn_xml = ElementTree.fromstring(_cn_response.content.decode('utf-8'))
                                        _cn_count = int(_cn_xml.find('tCnt').text)
                                        if _cn_count == -1:
                                            _is_warning = True
                                            _response_info = {
                                                'ECD': self._code + '_Run',
                                                'CLS': 'Abnormal Result',
                                                'DES': _cn_xml.find('ntceInfo').text,
                                                'LNO': 0,
                                                'DTA': [_remark]
                                            }
                                            break
                                        elif _cn_count == 0:
                                            if self._max_year == _year:
                                                _cn_lookup_data['CN'] = _cn_lookup_data.get('OCN')
                                                _cn_lookup_data['EN'] = None
                                                _cn_lookup_data['RD'] = None
                                                _cn_lookup_data['VS'] = '조회결과가 존재하지 않습니다.'
                                                _lookup_data_list.append(_cn_lookup_data.copy())
                                            else:
                                                continue
                                        else:
                                            for _cn_data in _cn_xml.findall('expFfmnBrkdCbnoQryRsltVo'):
                                                _cn_lookup_data['CN'] = _cn_data.findall('cbno')[0].text
                                                _cn_lookup_data['EN'] = _cn_data.findall('expDclrNo')[0].text
                                                _cn_lookup_data['RD'] = _cn_data.findall('dclrDttm')[0].text
                                                _cn_lookup_data['VS'] = _cn_data.findall('vhclPrgsStts')[0].text
                                                _lookup_data_list.append(_cn_lookup_data.copy())
                                            break
                                    else:
                                        _is_critical = True
                                        _remark['DESC'] = 'The Unipass API sent a {} code.'.format(_cn_response.status_code)
                                        _response_info = {
                                            'ECD': self._code + '_Run',
                                            'CLS': 'Non-200 Code',
                                            'DES': _cn_response.status_code,
                                            'LNO': 0,
                                            'DTA': [_remark]
                                        }
                                        break
                                else:
                                    break
                            for _en_lookup_data in _lookup_data_list:
                                _en_lookup_data['BL'] = []
                                _en_lookup_data['EC'] = 0
                                _en_response = None
                                if _en_lookup_data['EN'] is not None:
                                    _remark['LCN'] = _en_lookup_data['CN']
                                    _remark['LEN'] = _en_lookup_data['EN']
                                    _remark['URL'] = self._unipass_en_url.format(
                                        self._unipass_en_key, _en_lookup_data['EN']
                                    )
                                    for _ in range(11):
                                        try:
                                            _en_response = requests.get(_remark['URL'])
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
                                                _en_response = None
                                                break
                                            else:
                                                try:
                                                    time.sleep(2)
                                                    _en_response = requests.get(_remark['URL'])
                                                    break
                                                except requests.exceptions.ConnectionError:
                                                    _en_response = None
                                                    continue
                                    if not _is_critical:
                                        if _en_response.status_code == 200:
                                            _en_xml = ElementTree.fromstring(_en_response.content.decode('utf-8'))
                                            _en_count = int(_en_xml.find('tCnt').text)
                                            _en_lookup_data['EC'] = _en_count
                                            if _en_count == -1:
                                                _is_warning = True
                                                _response_info = {
                                                    'ECD': self._code + '_Run',
                                                    'CLS': 'Abnormal Result',
                                                    'DES': _en_xml.find('ntceInfo').text,
                                                    'LNO': 0,
                                                    'DTA': [_remark]
                                                }
                                            elif _en_count == 0:
                                                pass
                                            else:
                                                _en_data = _en_xml.findall('expDclrNoPrExpFfmnBrkdQryRsltVo')[0]
                                                _en_lookup_data['ES'] = _en_data.find('exppnConm').text
                                                _en_lookup_data['M'] = _en_data.find('mnurConm').text
                                                _en_lookup_data['LDD'] = str(_en_data.find('loadDtyTmlm').text)[:4] + '-' + str(_en_data.find('loadDtyTmlm').text)[4:6] + '-' + str(_en_data.find('loadDtyTmlm').text)[6:]
                                                _en_lookup_data['ACD'] = str(_en_data.find('acptDt').text)[:4] + '-' + str(_en_data.find('acptDt').text)[4:6] + '-' + str(_en_data.find('acptDt').text)[6:]
                                                _en_lookup_data['TA'] = int(_en_data.find('csclPckGcnt').text)
                                                _en_lookup_data['TW'] = int(float(_en_data.find('csclWght').text))
                                                _en_lookup_data['PS'] = _en_data.find('shpmCmplYn').text
                                                _en_lookup_data['SN'] = _en_data.find('sanm').text
                                                _en_lookup_data['PSA'] = int(_en_data.find('shpmPckGcnt').text)
                                                _en_lookup_data['PSW'] = int(_en_data.find('shpmWght').text)
                                                _en_lookup_data['ISP'] = _en_data.find('ldpInscTrgtYn').text
                                                _en_lookup_data['RA'] = _en_lookup_data['TA'] - _en_lookup_data['PSA']
                                                _en_lookup_data['RW'] = _en_lookup_data['TW'] - _en_lookup_data['PSW']
                                                if _en_count != 0:
                                                    for _bl_data in _en_xml.findall('expDclrNoPrExpFfmnBrkdDtlQryRsltVo'):
                                                        _en_lookup_data['BL'].append({
                                                            'NO': _bl_data.findall('blNo')[0].text,
                                                            'SD': str(_bl_data.findall('tkofDt')[0].text)[:4] + '-' + str(_bl_data.findall('tkofDt')[0].text)[4:6] + '-' + str(_bl_data.findall('tkofDt')[0].text)[6:],
                                                            'PSA': int(_bl_data.findall('shpmPckGcnt')[0].text),
                                                            'PLW': int(_bl_data.findall('shpmWght')[0].text)
                                                        })
                                        else:
                                            _is_critical = True
                                            _remark['DESC'] = 'The Unipass API sent a {} code.'.format(_en_response.status_code)
                                            _response_info = {
                                                'ECD': self._code + '_Run',
                                                'CLS': 'Non-200 Code',
                                                'DES': _en_response.status_code,
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
                                        _remark['OCN'] = _lookup_data['OCN']
                                        _remark['LCN'] = _lookup_data['CN']
                                        _remark['LEN'] = _lookup_data['EN']
                                        _remark.__setitem__('DESC', None)
                                        _remark.__setitem__('URL', None)

                                        _query = 'INSERT INTO dict_chassis_number(chassis_number) VALUES(\'{}\')'.format(_lookup_data['CN'])
                                        _der = self._dbm.execute_query(_query)
                                        if _der[0]:
                                            _d_cn_index = _der[1]
                                        else:
                                            if 'Duplicate entry' in str(_der[2]['DES']):
                                                _query2 = 'SELECT d_chassis_index FROM dict_chassis_number WHERE chassis_number = \'{}\''.format(_lookup_data['CN'])
                                                _der2 = self._dbm.execute_query(_query2)
                                                if _der2[0]:
                                                    _d_cn_index = _der2[1][0][0]
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
                                        else:
                                            _d_en_index = 0
                                        _query = '''
                                        INSERT INTO export_info_chassis_number(d_export_index, d_chassis_index, report_date, vehicle_status)
                                        VALUES({}, {}, {}, '{}')
                                        '''.format(
                                            _d_en_index,
                                            _d_cn_index,
                                            '\'{}\''.format(_lookup_data['RD']) if _lookup_data['RD'] is not None else '\'1000-01-01\'',
                                            _lookup_data['VS']
                                        )
                                        _der = self._dbm.execute_query(_query)
                                        if _der[0]:
                                            _ei_chassis_index = _der[1]
                                        else:
                                            _der[2]['DTA'].append(_remark.copy())
                                            self.send(True, False, _der[2])
                                            break
                                        if _lookup_data['EC'] > 0:
                                            _query = '''
                                            INSERT INTO export_info_export_number(
                                                d_export_index,
                                                export_shipper,
                                                maker,
                                                load_duty_deadline,
                                                accept_date,
                                                total_amount,
                                                total_weight,
                                                flag_pre_shipping,
                                                ship_name,
                                                pre_shipping_amount,
                                                pre_shipping_weight,
                                                flag_inspection,
                                                remain_amount,
                                                remain_weight
                                            )
                                            VALUES(
                                                {},
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
                                                {},
                                                {}
                                            )
                                            '''.format(
                                                _d_en_index,
                                                _lookup_data['ES'],
                                                _lookup_data['M'],
                                                _lookup_data['LDD'],
                                                _lookup_data['ACD'],
                                                _lookup_data['TA'],
                                                _lookup_data['TW'],
                                                _lookup_data['PS'],
                                                _lookup_data['SN'],
                                                _lookup_data['PSA'],
                                                _lookup_data['PSW'],
                                                _lookup_data['ISP'],
                                                _lookup_data['RA'],
                                                _lookup_data['RW']
                                            )
                                            _der = self._dbm.execute_query(_query)
                                            if _der[0]:
                                                _ei_export_index = _der[1]
                                            else:
                                                _der[2]['DTA'].append(_remark.copy())
                                                self.send(True, False, _der[2])
                                                break
                                        else:
                                            _ei_export_index = 0
                                        _query = '''
                                        INSERT INTO lookup_chassis_result(l_chassis_index, ei_chassis_index, ei_export_index, excel_row)
                                        VALUES({}, {}, {}, {})
                                        '''.format(self._inquiry_index, _ei_chassis_index, _ei_export_index, _lookup_data['ERN'])
                                        _der = self._dbm.execute_query(_query)
                                        if _der[0]:
                                            _r_chassis_index = _der[1]
                                        else:
                                            _der[2]['DTA'].append(_remark.copy())
                                            self.send(True, False, _der[2])
                                            break
                                        if len(_lookup_data['BL']) != 0:
                                            for _bl in _lookup_data['BL']:
                                                if _bl.get('NO') is not None:
                                                    _query = 'INSERT INTO dict_bl_number(bl_number) VALUES(\'{}\')'.format(_bl.get('NO'))
                                                    _der = self._dbm.execute_query(_query)
                                                    if _der[0]:
                                                        _d_bl_index = _der[1]
                                                    else:
                                                        if 'Duplicate entry' in str(_der[2]['DES']):
                                                            _query2 = 'SELECT d_bl_index FROM dict_bl_number WHERE bl_number = \'{}\''.format(_bl.get('NO'))
                                                            _der2 = self._dbm.execute_query(_query2)
                                                            if _der2[0]:
                                                                _d_bl_index = _der2[1][0][0]
                                                            else:
                                                                _der2[2]['DTA'].append(_remark.copy())
                                                                self.send(True, False, _der2[2])
                                                                break
                                                        else:
                                                            _der[2]['DTA'].append(_remark.copy())
                                                            self.send(True, False, _der[2])
                                                            break
                                                    _query = '''INSERT INTO export_info_export_number_bl(ei_export_index, d_bl_index, shipment_date, pre_shipping_amount, pre_shipping_weight)
                                                    VALUES({}, '{}', '{}', {}, {})'''.format(
                                                        _ei_export_index,
                                                        _d_bl_index,
                                                        _bl.get('SD'),
                                                        _bl.get('PSA'),
                                                        _bl.get('PLW')
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
            self.send(False, True, None)


class GetLookUpDataList(_ThreadModel):
    def __init__(self, _thread_queue: Queue, _manager_index: int, _inquiry_index: int, _order_str: str, _app_queue: [Queue, None] = None):
        self._code, self._manager_index, self._inquiry_index, self._order_str = 'ModelVehicleLookup_GetLookUpDataList', _manager_index, _inquiry_index, _order_str
        self._dbm = _DataBaseManager(self._code)
        super().__init__(_thread_queue, _app_queue)

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
                    self._inquiry_index,
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
