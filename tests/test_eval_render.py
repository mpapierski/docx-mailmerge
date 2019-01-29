import unittest
import tempfile
from os import path
from lxml import etree
from decimal import Decimal
import subprocess

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin
from tests.utils import has_soffice


@unittest.skipIf(has_soffice() is None, 'soffice executable not found in $PATH')
class EvalRenderTest(EtreeMixin, unittest.TestCase):
    def test_man(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_eval.docx')) as document:
            self.assertEqual(document.get_merge_fields(), {'Price', 'Missing'})

            document.merge(**{'Price': Decimal('1234.56')})

            with tempfile.NamedTemporaryFile() as outfile:
                document.write(outfile)
                outfile.flush()
                stdout = subprocess.check_output(['soffice', '--headless', '--cat', outfile.name])
                stdout = stdout.decode('utf-8')
                if stdout.startswith('\ufeff'):
                    stdout = stdout[1:]
                self.assertEqual(stdout.splitlines(), [
                    '1\xa0234,56 zł',
                    '1\xa0234,56 zł',
                    '1\xa0234,56 PLN',
                    '1\xa0234,56 PLN',
                    '1\xa0234,56 złotego polskiego',
                    '1\xa0234,56 złotego polskiego',
                    '',
                    ''
                ])
