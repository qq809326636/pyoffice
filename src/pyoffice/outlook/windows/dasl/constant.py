__all__ = ['OperatorEnum']

LinkEnum = {
    'and': {
        'code': 10,
        'op': 'AND'
    },
    'or': {
        'code': 11,
        'op': 'OR'
    }
}

OperatorEnum = {
    'root': {
        'code': -1,
        'op': ''
    },
    'eq': {
        'code': 10,
        'op': '='
    },
    'ne': {
        'code': 11,
        'op': '<>'
    },
    'lt': {
        'code': 20,
        'op': '<'
    },
    'le': {
        'code': 21,
        'op': '<='
    },
    'gt': {
        'code': 31,
        'op': '>'
    },
    'ge': {
        'code': 40,
        'op': '>='
    }
}
