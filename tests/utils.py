from functools import lru_cache
import subprocess


class EtreeMixin(object):
    def assert_equal_tree(self, lhs, rhs):
        """
        Compares two instances of ElementTree are equivalent.
        """
        self.assertEqual(lhs.tag, rhs.tag)
        self.assertEqual(len(lhs), len(rhs))
        self.assertEqual(lhs.text or '', rhs.text or '')
        self.assertEqual(sorted(lhs.items()), sorted(rhs.items()))
        for lhs_child, rhs_child in zip(lhs, rhs):
            self.assert_equal_tree(lhs_child, rhs_child)


def get_document_body_part(document):
    for part in document.parts.values():
        if part.getroot().tag.endswith('}document'):
            return part

    raise AssertionError("main document body not found in document.parts")


@lru_cache()
def has_soffice():
    try:
        stdout = subprocess.check_output(['which', 'soffice'])
        stdout = stdout.decode('utf-8')
        lines = stdout.splitlines()
        return lines[0]
    except subprocess.CalledProcessError:
        return None
