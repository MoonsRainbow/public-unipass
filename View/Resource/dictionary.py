HEAD_CATEGORY = {
    'No': {
        't': 'No',
        'w': 40,
        'v': ['PS', 'S', 'L', 'H', 'HP', 'A'],
        'e': True
    },
    'ID': {
        't': '아이디',
        'w': 180,
        'v': ['H', 'A'],
        'e': False
    },
    'ManagerName': {
        't': '이름',
        'w': 180,
        'v': ['A'],
        'e': False
    },
    'Result': {
        't': '결과',
        'w': 60,
        'v': ['S', 'HP'],
        'e': False
    },
    'ExlRowNo': {
        't': 'Excel No',
        'w': 80,
        'v': ['S', 'HP'],
        'e': True
    },
    'ExportDeclarationNumber': {
        't': '수출신고번호',
        'w': 180,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'ChassisNo': {
        't': '차대번호',
        'w': 180,
        'v': ['PS', 'S', 'L', 'HP'],
        'e': True
    },
    'InquiryDate': {
        't': '조회일자',
        'w': 100,
        'v': ['L'],
        'e': False
    },
    'InquiryDateForHistory': {
        't': '조회일자',
        'w': 200,
        'v': ['H'],
        'e': False
    },
    'InquiryCount': {
        't': '조회개수',
        'w': 80,
        'v': ['H'],
        'e': False
    },
    'ReportDate': {
        't': '신고일자',
        'w': 140,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'LoadDutyDeadline': {
        't': '적재의무기한',
        'w': 140,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'VehicleStatus': {
        't': '차량진행상태',
        'w': 450,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'TotalAmount': {
        't': '총 개수',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'TotalWeight': {
        't': '총 중량',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'ShippingAmount': {
        't': '선적 개수',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'ShippingWeight': {
        't': '선적 중량',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'RemainAmount': {
        't': '잔여 개수',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'RemainWeight': {
        't': '잔여 중량',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'Inspection': {
        't': '검사 유무',
        'w': 80,
        'v': ['S', 'L', 'HP'],
        'e': True
    },
    'Extra': {
        't': '비고',
        'w': 400,
        'v': ['PS'],
        'e': False
    }
}

VEHICLE_LOOKUP_RESULT = 'VLR'
VEHICLE_LIST = 'VL'
LOOKUP_HISTORY_VEHICLE_LIST = 'LHVL'
VEHICLE_HEAD_CATEGORY = {
    'KEY_INDEX': {
        'TEXT': 'No',
        'WIDTH': 40,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'RESULT': {
        'TEXT': '결과',
        'WIDTH': 60,
        'TABLE': [VEHICLE_LOOKUP_RESULT, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': False
    },
    'EXCEL_ROW_NO': {
        'TEXT': '엑셀 행',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'ORIGIN_EXPORT_NUMBER': {
        'TEXT': '수출신고번호',
        'WIDTH': 180,
        'TABLE': [],
        'IN_EXCEL': True
    },
    'FORMAT_EXPORT_NUMBER': {
        'TEXT': '수출신고번호',
        'WIDTH': 180,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': False
    },
    'CHASSIS_NUMBER': {
        'TEXT': '차대번호',
        'WIDTH': 180,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'LOOKUP_DATE': {
        'TEXT': '조회일자',
        'WIDTH': 100,
        'TABLE': [VEHICLE_LIST],
        'IN_EXCEL': False
    },
    'REPORT_DATE': {
        'TEXT': '신고일자',
        'WIDTH': 140,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'LOAD_DUTY_DEADLINE': {
        'TEXT': '적재의무기한',
        'WIDTH': 140,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'VEHICLE_STATUS': {
        'TEXT': '차량진행상태',
        'WIDTH': 450,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'TOTAL_AMOUNT': {
        'TEXT': '총 개수',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'TOTAL_WEIGHT': {
        'TEXT': '총 중량',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'PRE_AMOUNT': {
        'TEXT': '선적 개수',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'PRE_WEIGHT': {
        'TEXT': '선적 중량',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'REMAIN_AMOUNT': {
        'TEXT': '잔여 개수',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'REMAIN_WEIGHT': {
        'TEXT': '잔여 중량',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    },
    'FLAG_INSPECTION': {
        'TEXT': '검사 유무',
        'WIDTH': 80,
        'TABLE': [VEHICLE_LOOKUP_RESULT, VEHICLE_LIST, LOOKUP_HISTORY_VEHICLE_LIST],
        'IN_EXCEL': True
    }
}

BL_LIST = 'BL'
LOOKUP_HISTORY_BL_LIST = 'LHBL'
BL_HEAD_CATEGORY = {
    'KEY_INDEX': {
        'TEXT': 'No',
        'WIDTH': 40,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'EXCEL_ROW_NO': {
        'TEXT': '엑셀 행',
        'WIDTH': 80,
        'TABLE': [LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'BL_NUMBER': {
        'TEXT': 'B/L 번호',
        'WIDTH': 180,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'ORIGIN_EXPORT_NUMBER': {
        'TEXT': '수출신고번호',
        'WIDTH': 180,
        'TABLE': [],
        'IN_EXCEL': True
    },
    'FORMAT_EXPORT_NUMBER': {
        'TEXT': '수출신고번호',
        'WIDTH': 180,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': False
    },
    'MANAGEMENT_NUMBER': {
        'TEXT': '적하목록관리번호',
        'WIDTH': 180,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'LOOKUP_DATE': {
        'TEXT': '조회일자',
        'WIDTH': 100,
        'TABLE': [BL_LIST],
        'IN_EXCEL': True
    },
    'ACCEPT_DATE': {
        'TEXT': '수리일자',
        'WIDTH': 100,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'LOAD_DUTY_DEADLINE': {
        'TEXT': '적재의무기한',
        'WIDTH': 100,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'EXPORT_SHIPPER': {
        'TEXT': '수출자',
        'WIDTH': 180,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'TOTAL_AMOUNT': {
        'TEXT': '총 개수',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'TOTAL_WEIGHT': {
        'TEXT': '총 중량',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'SHIPMENT_PLACE': {
        'TEXT': '선기적지',
        'WIDTH': 100,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'DEPARTURE_DATE': {
        'TEXT': '출항일자',
        'WIDTH': 100,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'PRE_AMOUNT': {
        'TEXT': '선적 개수',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'PRE_WEIGHT': {
        'TEXT': '선적 중량',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'DIVISION_COLLECT': {
        'TEXT': '분할회수',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    },
    'FLAG_PRE_SHIPPING': {
        'TEXT': '선적완료여부',
        'WIDTH': 80,
        'TABLE': [BL_LIST, LOOKUP_HISTORY_BL_LIST],
        'IN_EXCEL': True
    }
}


VEHICLE_LOOKUP_BEFORE = 'VLB'
LOOKUP_HISTORY_LIST = 'LHL'
ADMIN_SETTING_MANAGER_LIST = 'ASML'
COMMON_HEAD_CATEGORY = {
    'KEY_INDEX': {
        'TEXT': 'No',
        'WIDTH': 40,
        'TABLE': [VEHICLE_LOOKUP_BEFORE, LOOKUP_HISTORY_LIST, ADMIN_SETTING_MANAGER_LIST],
        'IN_EXCEL': False
    },
    'MANAGER_ID': {
        'TEXT': '아이디',
        'WIDTH': 180,
        'TABLE': [LOOKUP_HISTORY_LIST, ADMIN_SETTING_MANAGER_LIST],
        'IN_EXCEL': False
    },
    'MANAGER_NAME': {
        'TEXT': '이름',
        'WIDTH': 180,
        'TABLE': [ADMIN_SETTING_MANAGER_LIST],
        'IN_EXCEL': False
    },
    'CHASSIS_NUMBER': {
        'TEXT': '차대번호',
        'WIDTH': 180,
        'TABLE': [VEHICLE_LOOKUP_BEFORE]
    },
    'LOOKUP_DATE': {
        'TEXT': '조회일자',
        'WIDTH': 200,
        'TABLE': [LOOKUP_HISTORY_LIST],
        'IN_EXCEL': False
    },
    'LOOKUP_COUNT': {
        'TEXT': '조회개수',
        'WIDTH': 80,
        'TABLE': [LOOKUP_HISTORY_LIST],
        'IN_EXCEL': False
    },
    'EXTRA': {
        'TEXT': '비고',
        'WIDTH': 400,
        'TABLE': [VEHICLE_LOOKUP_BEFORE],
        'IN_EXCEL': False
    }
}
