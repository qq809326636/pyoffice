

__all__ = ['EscapingEnum',
           'DASLPrefix']


class DASLPrefix:
    PREFIX = '@SQL='


class EscapingEnum:
    SPACE_CHARACTER = '%20'
    Double_quote = '%22'
    Single_quote = '%27'
    Percent_character = '%25'
