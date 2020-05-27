from .DASLDate import DASLDate

__all__ = ['OperatorEnum',
           'LinkEnum',
           'PropertyEnum',
           'EscapingEnum',
           'DASLOperatorEnum',
           'DASLPrefix']


class DASLPrefix:
    PREFIX = '@SQL='


class DASLOperatorEnum:
    STARTS_WITH = 'ci_startswith'
    PHRASE_MATCH = 'ci_phrasematch'
    LIKE = 'like'


class EscapingEnum:
    SPACE_CHARACTER = '%20'
    Double_quote = '%22'
    Single_quote = '%27'
    Percent_character = '%25'


PropertyEnum = {
    'sender': {
        'ref': 'fromname',
        'parent': 'httpmail',
        'type': str
    },
    'recipient': {
        'ref': 'displayto',
        'parent': 'httpmail',
        'type': str
    },
    'cc': {
        'ref': 'displaycc',
        'parent': 'httpmail',
        'type': str
    },
    'bcc': {
        'ref': 'resources',
        'parent': 'calendar',
        'type': str
    },
    'sentDate': {
        'ref': 'date',
        'parent': 'httpmail',
        'type': DASLDate
    },
    'subject': {
        'ref': 'subject',
        'parent': 'httpmail',
        'type': str
    },
    'message': {
        'ref': 'textdescription',
        'parent': 'httpmail',
        'type': str
    },
    'importance': {
        'ref': 'importance',
        'parent': 'httpmail',
        'type': int
    },
    'attachment': {
        'ref': 'hasattachment',
        'parent': 'httpmail',
        'type': bool
    },
    'read': {
        'ref': 'read',
        'parent': 'httpmail',
        'type': bool
    },

    'created': {
        'ref': 'created',
        'parent': 'httpmail',
        'type': str
    },
    'receivedDate': {
        'ref': 'read',
        'parent': 'httpmail',
        'type': DASLDate
    },
    'account': {
        'ref': 'account',
        'parent': 'contacts',
        'type': str
    }
}

LinkEnum = {
    'and': {
        'code': 10,
        'op': 'AND'
    },
    'or': {
        'code': 11,
        'op': 'OR'
    },
    'root': {
        'code': -1,
        'op': ''
    }
}

OperatorEnum = {
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
        'code': 30,
        'op': '>'
    },
    'ge': {
        'code': 31,
        'op': '>='
    },
    'contains': {
        'code': 40
    },
    'startswith': {
        'code': 41
    },
    'endswith': {
        'code': 42
    }
}
