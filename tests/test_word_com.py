import copy
import sys
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from word_com import text_hash, validate_request

class RequestTests(unittest.TestCase):
    def setUp(self):
        self.target = {'pid': 1, 'hwnd': 2, 'name': 'test', 'full_name': 'test'}
        self.req = {'target': self.target, 'text_sha256': text_hash('abc\r'), 'operation': {'kind': 'replace', 'start': 0, 'end': 3, 'expected': 'abc', 'text': 'xyz'}}
    def test_valid(self):
        self.assertEqual(validate_request(self.req, self.target, 'abc\r')['kind'], 'replace')
    def test_stale(self):
        with self.assertRaises(ValueError): validate_request(self.req, self.target, 'abcd\r')
    def test_identity(self):
        with self.assertRaises(ValueError): validate_request(self.req, {}, 'abc\r')
    def test_bad_requests(self):
        for patch in ({'start': -1}, {'kind': 'evil'}, {'expected': 'abc\r'}, {'text': '\x07'}, {'kind': 'insert'}, {'kind': 'format', 'font': {'bold': 'yes'}}):
            req = copy.deepcopy(self.req)
            req['operation'].update(patch)
            with self.assertRaises(ValueError): validate_request(req, self.target, 'abc\r')
    def test_unicode_hash(self):
        self.assertEqual(len(text_hash('中文😀\r')), 64)

if __name__ == '__main__': unittest.main()
