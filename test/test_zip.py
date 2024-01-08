#pylint: skip-file
"""unit tests for features"""
import tempfile
import unittest
from inspect import getsourcefile as _getsourcefile
import os.path as _path
import funclite.iolib as iolib

import docs
import docs.zip as docs_zip

class Test(unittest.TestCase):
    """unittest for keypoints"""

    def setUp(self):
        """setup variables etc for use in test cases
        """
        self.this_file_path = _path.normpath(iolib.get_file_parts2(_path.abspath(_getsourcefile(lambda: 0)))[0])
        self.module_root = _path.normpath(tuple(docs.__path__)[0]) #root
        self.zip_file_path = _path.normpath(_path.join(self.module_root, 'test/bin'))
        self.zip_file = _path.normpath(self.zip_file_path + '/root.zip')
        self.temp_dir = tempfile.gettempdir()


    #@unittest.skip("Temporaily disabled while debugging")
    def test_extract(self):
        zipto = _path.normpath(_path.join(self.temp_dir))
        docs_zip.extract(self.zip_file, zipto, match_folder_name=('1',), match_file_name=('1',), match_file_ext=('.docx',)) #extract folder matching 1 and files matching 1
        docs_zip.extract(self.zip_file, zipto) #full extract
        pass


if __name__ == '__main__':
    unittest.main(verbosity=2)
