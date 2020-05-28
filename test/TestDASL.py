from pyoffice.outlook.windows import *
import pytest
import datetime


class TestDASL:

    @pytest.fixture(scope='module')
    def builder(self):
        return DASLBuilder()

    @pytest.fixture(scope='module')
    def senderCondition(self):
        cond = DASLCondition()
        cond.prop = 'sender'
        cond.op = 40
        cond.val = '1data'
        return cond

    @pytest.fixture(scope='module')
    def recipientCondition(self):
        cond = DASLCondition()
        cond.prop = 'recipient'
        cond.op = 40
        cond.val = '1data'
        return cond

    @pytest.fixture(scope='module')
    def ccCondition(self):
        cond = DASLCondition()
        cond.prop = 'cc'
        cond.op = 40
        cond.val = '1data'
        return cond

    @pytest.fixture(scope='module')
    def bccCondition(self):
        cond = DASLCondition()
        cond.prop = 'bcc'
        cond.op = 40
        cond.val = '1data'
        return cond

    @pytest.fixture(scope='module')
    def sentDateCondition(self):
        cond = DASLCondition()
        cond.prop = 'sentDate'
        cond.op = 20
        cond.val = datetime.datetime.now()
        return cond

    @pytest.fixture(scope='module')
    def subjectCondition(self):
        cond = DASLCondition()
        cond.prop = 'subject'
        cond.op = 40
        cond.val = 'test'
        return cond

    @pytest.fixture(scope='module')
    def messageCondition(self):
        cond = DASLCondition()
        cond.prop = 'message'
        cond.op = 40
        cond.val = '1data'
        return cond

    @pytest.fixture(scope='module')
    def importanceCondition(self):
        cond = DASLCondition()
        cond.prop = 'importance'
        cond.op = 10
        cond.val = '1'
        return cond

    @pytest.fixture(scope='module')
    def attachmentCondition(self):
        cond = DASLCondition()
        cond.prop = 'attachment'
        cond.op = 10
        cond.val = '1'
        return cond

    @pytest.fixture(scope='module')
    def readCondition(self):
        cond = DASLCondition()
        cond.prop = 'read'
        cond.op = 10
        cond.val = '1'
        return cond

    def test_builder(self,
                     builder,
                     senderCondition,
                     recipientCondition,
                     ccCondition,
                     bccCondition,
                     sentDateCondition,
                     subjectCondition,
                     messageCondition,
                     importanceCondition,
                     attachmentCondition,
                     readCondition):
        print()

        #
        builder.add(senderCondition)

        #
        # recipientCondition.linker = 10
        # builder.add(recipientCondition)
        #
        # #
        # ccCondition.linker = 10
        # builder.add(ccCondition)
        #
        # #
        # bccCondition.linker = 10
        # builder.add(bccCondition)

        #
        # sentDateCondition.linker = 10
        # builder.add(sentDateCondition)

        #
        # importanceCondition.linker = 10
        # builder.add(importanceCondition)

        #
        attachmentCondition.link = 10
        builder.add(attachmentCondition)

        #
        readCondition.link = 10
        builder.add(readCondition)

        print(f'[INFO]: DASL is "{builder.build()}"')
