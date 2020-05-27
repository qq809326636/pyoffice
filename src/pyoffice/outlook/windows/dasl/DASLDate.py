import datetime

__all__ = ['DASLDate']


class DASLDate:

    def __init__(self,
                 val):
        self._val = val

        if isinstance(self._val, datetime.datetime):
            self._date = self._val
        elif isinstance(self._val, datetime.time):
            self._date = self._val
        else:
            self._date = datetime.datetime.fromisoformat(self._val)
            # self._date = datetime.datetime.strptime(self._val)

    def toDatetime(self):
        return self._date

    def toString(self):
        return datetime.datetime.strftime(self._val, '%Y-%m-%d %H:%M:%S')

    def __str__(self):
        return self.toString()

    def __repr__(self):
        return self.toString()
