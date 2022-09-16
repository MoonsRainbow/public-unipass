import sys
import queue
import pymysql
import threading


class _ThreadModel(threading.Thread):
    """
    비동기, 동기 처리 서브 스레드 클래스.

    Args:
        _thread_queue (queue.Queue): App으로 Data를 발신하기 위한  Queue Instance
        _app_queue (queue.Queue): App에서 Data를 수신받기 위한 Queue Instance

    Attributes:
        _thread_queue (queue.Queue):
        _app_queue (queue.Queue):
    """
    def __init__(self, _thread_queue: queue.Queue, _app_queue: [queue.Queue, None]):
        self._thread_queue, self._app_queue = _thread_queue, _app_queue
        threading.Thread.__init__(self)

    def send(self, _run: bool, _success: bool, _data: [any, None] = None):
        """
        서브 스레드 데이터를 메인 스레드로 전달하는 메소드.

        Args:
            _run (bool): 작동 상태 플래그.
            _success (bool): 처리 성공 여부 플래그.
            _data (bool, None): 데이터
        """
        self._thread_queue.put(
            {
                'RUN': _run,
                'SUCCESS': _success,
                'DATA': _data
            }
        )

    def recv(self):
        if self._app_queue is not None:
            if self._app_queue.empty():
                return None
            else:
                try:
                    return self._app_queue.get(block=False)
                except queue.Empty:
                    return None
        else:
            return None


class _DataBaseManager:
    """
    데이터 베이스 연결 관리 클래스.

    Attributes:
        _code (str): 클래스 고유 코드
        _conn (pymysql.connect, None):
        _info (dict): 데이터 베이스 정보
    """
    def __init__(self, _code):
        self._code = _code + '_DBM'
        self._conn: [pymysql.connect, None] = None
        self._info: dict = {}

    def connecting(self):
        """
        데이터 베이스 연결 메소드.

        Returns:
            bool, {} or {'ECD', 'CLS', 'DES', 'LNO', 'DTA'}
        """
        try:
            self._conn = pymysql.connect(
                host=self._info['HOST'],
                port=self._info['PORT'],
                user=self._info['USER'],
                password=self._info['PW'],
                database=self._info['DB']
            )
            return True, {}
        except:
            self._conn = None
            return False, {
                'ECD': self._code + '_Connecting',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': []
            }

    def disconnecting(self):
        """
        데이터 베이스 종료 메소드.
        """
        try:
            self._conn.close()
            self._conn = None
        except:
            pass

    def execute_query(self, sql: str):
        """
        데이터 베이스 CRUD 처리 메소드.

        Args:
            sql (str): SQL Query

        Returns:
            bool, int or tuple or none, {} or {'ECD', 'CLS', 'DES', 'LNO', 'DTA'}
        """
        sql = sql.replace("\n", " ").strip()
        try:
            _result = None
            with self._conn.cursor() as _conn:
                _conn.execute(sql)
                self._conn.commit()
                if sql.lower().strip().startswith("insert", 0):
                    _result = _conn.lastrowid
                else:
                    _result = _conn.fetchall()
            return True, _result, {}
        except:
            return False, None, {
                'ECD': self._code + '_ExecuteQuery',
                'CLS': sys.exc_info()[0],
                'DES': sys.exc_info()[1],
                'LNO': sys.exc_info()[2].tb_lineno,
                'DTA': [sql]
            }
