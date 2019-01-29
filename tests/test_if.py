import unittest
import tempfile
from os import path
from lxml import etree
import subprocess

from mailmerge import MailMerge, NAMESPACES
from tests.utils import EtreeMixin
from tests.utils import has_soffice


@unittest.skipIf(has_soffice() is None, 'soffice executable not found in $PATH')
class IfTest(EtreeMixin, unittest.TestCase):
    def test_man(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_if.docx')) as document:
            self.assertEqual(document.get_merge_fields(), set())

            document.merge(**{'Gender': 'Man'})

            with tempfile.NamedTemporaryFile() as outfile:
                document.write(outfile)
                outfile.flush()
                stdout = subprocess.check_output(['soffice', '--headless', '--cat', outfile.name])
                stdout = stdout.decode('utf-8')
                if stdout.startswith('\ufeff'):
                    stdout = stdout[1:]
                self.assertEqual(stdout.splitlines(), ['True', 'False'])

    def test_woman(self):
        with MailMerge(path.join(path.dirname(__file__), 'test_if.docx')) as document:
            self.assertEqual(document.get_merge_fields(), set())

            document.merge(**{'Gender': 'Woman'})

            with tempfile.NamedTemporaryFile() as outfile:
                document.write(outfile)
                outfile.flush()
                stdout = subprocess.check_output(['soffice', '--headless', '--cat', outfile.name])
                stdout = stdout.decode('utf-8')
                if stdout.startswith('\ufeff'):
                    stdout = stdout[1:]
                self.assertEqual(stdout.splitlines(), ['False', 'True'])
