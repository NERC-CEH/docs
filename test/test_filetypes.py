#pylint: skip-file
"""unit tests for features"""
import unittest
from inspect import getsourcefile as _getsourcefile
import os.path as _path
import funclite.iolib as iolib
import docs
import docs.filetypes as ft

class Test(unittest.TestCase):
    """unittest for keypoints"""

    def setUp(self):
        """setup variables etc for use in test cases
        """
        self.this_file_path = _path.normpath(iolib.get_file_parts2(_path.abspath(_getsourcefile(lambda: 0)))[0])
        self.module_root = _path.normpath(tuple(docs.__path__)[0]) #root




    #@unittest.skip("Temporaily disabled while debugging")
    def test_it(self):
        S = ft.Images.dotted()
        S = ft.Images.asterixed()




if __name__ == '__main__':
    unittest.main(verbosity=2)
