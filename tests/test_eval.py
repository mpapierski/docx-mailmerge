from unittest import TestCase
from datetime import datetime
from mailmerge import MailMerge
import shlex


def code(text):
    return shlex.split(text, posix=False)


class TestEval(TestCase):

    def setUp(self):
        self.today = datetime(2018, 7, 1, 10, 12, 54, 0)

    def test_year(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'y'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'Y'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yy'), '18')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YY'), '18')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yyyy'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YYYY'), '2018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'yyyyy'), '20182018')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'YYYYY'), '20182018')

    def test_month(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'M'), '7')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'MM'), '07')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'MMMM'), self.today.strftime('%B'))

    def test_days(self):
        self.assertEqual(MailMerge.eval_strftime(self.today, 'd'), '1')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'dd'), '01')
        self.assertEqual(MailMerge.eval_strftime(self.today, 'dddd'), self.today.strftime('%A'))

    def test_strftime_none(self):
        self.assertEqual(MailMerge.eval_strftime(None, 'd'), '')

    def test_strftime_str(self):
        self.assertEqual(MailMerge.eval_strftime('foo', 'd'), 'foo')

    def test_eval(self):
        self.assertEqual(MailMerge.eval_star(self.today, ''), self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval_star(self.today.date(), ''), self.today.strftime('%x'))
        self.assertEqual(MailMerge.eval_star(self.today.time(), ''), self.today.strftime('%X'))

    def test_eval_none(self):
        self.assertEqual(MailMerge.eval(None, code('MERGEFIELD Foo \\@ "y-MM-dd" MERGEFORMAT')), '')

    def test_parse_code(self):
        self.assertEqual(MailMerge.eval(self.today, code('MERGEFIELD Foo \\* MERGEFORMAT')),
                         self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval(self.today, code('MERGEFIELD Foo')), self.today.strftime('%x %X'))
        self.assertEqual(MailMerge.eval(self.today, code('MERGEFIELD Foo \\@ "y-MM-dd" MERGEFORMAT')), '2018-07-01')

    def test_parse_str_formatter(self):
        self.assertEqual(MailMerge.eval('HeLlO WoRlD', code('MERGEFIELD Foo \\* Upper MERGEFORMAT')), 'HELLO WORLD')
        self.assertEqual(MailMerge.eval('HELLO WORLD', code('MERGEFIELD Foo \\* Lower MERGEFORMAT')), 'hello world')
        self.assertEqual(MailMerge.eval('HeLlO WoRlD', code(
            'MERGEFIELD Foo \\* Lower \\* Upper MERGEFORMAT')), 'HELLO WORLD')
        self.assertEqual(MailMerge.eval('HeLlO WoRlD', code('MERGEFIELD Foo \\* FirstCap MERGEFORMAT')), 'Hello world')
        self.assertEqual(MailMerge.eval('HeLlO WoRlD', code('MERGEFIELD Foo \\* Caps MERGEFORMAT')), 'Hello World')

    def test_parse_chained(self):
        self.assertEqual(MailMerge.eval(self.today, code('MERGEFIELD Foo \\* Upper')),
                         self.today.strftime('%x %X').upper())
        self.assertEqual(MailMerge.eval(self.today, code('MERGEFIELD Foo \\@ "MMMM" \\* Upper')),
                         self.today.strftime('%B').upper())

    def test_eval_if(self):
        # Simple equality
        # self.assertEqual(MailMerge.eval())
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo = "Bar" "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 'Bar'}), 'Yes')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo <> "Bar" "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 'Bar'}), 'No')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo <> 123 "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 123}), 'No')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo > 123 "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 123}), 'No')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo >= 123 "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 123}), 'Yes')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo < 123 "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 123}), 'No')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo <= 123 "Yes" "No" \\* MERGEFORMAT'), context={'Foo': 123}), 'Yes')

        self.assertEqual(MailMerge.eval('', code('IF Gender = "M*" "Yes" "No" \\* MERGEFORMAT'),
                                        context={'Gender': 'Man'}), 'Yes')
        self.assertEqual(MailMerge.eval('', code('IF Gender <> "M*" "Yes" "No" \\* MERGEFORMAT'),
                                        context={'Gender': 'Man'}), 'No')
        self.assertEqual(MailMerge.eval('', code('IF Gender = "?oma*" "Yes" "No" \\* MERGEFORMAT'),
                                        context={'Gender': 'Woman'}), 'Yes')
        self.assertEqual(MailMerge.eval('', code('IF Gender = "?oma*" "Yes" "No" \\* MERGEFORMAT'),
                                        context={'Gender': 'Xom'}), 'No')
        self.assertEqual(MailMerge.eval('', code('IF Gender = "?oma*" "Yes" "No" \\* MERGEFORMAT'),
                                        context={'Gender': 'Xomannnnn'}), 'Yes')

        self.assertEqual(MailMerge.eval(
            '', code('IF Foo = Foo "This is a long answer" "Nope" \\* MERGEFORMAT'), context={}), 'This is a long answer')
        self.assertEqual(MailMerge.eval(
            '', code('IF Foo <> Foo "This is a long answer" "N o p e" \\* MERGEFORMAT'), context={}), 'N o p e')
