#pylint: skip-file
"""unit tests for features"""
import unittest
from inspect import getsourcefile as _getsourcefile
import os.path as _path
import funclite.iolib as iolib
import img2pdf
import docs
import docs.topdf as topdf


class Test(unittest.TestCase):
    """unittest for keypoints"""

    def setUp(self):
        """setup variables etc for use in test cases
        """
        self.this_file_path = _path.normpath(iolib.get_file_parts2(_path.abspath(_getsourcefile(lambda: 0)))[0])
        self.module_root = _path.normpath(tuple(docs.__path__)[0]) #root
        self.test_images_path = _path.normpath(_path.join(self.module_root, 'test/bin/images'))
        self.x1_image_path = _path.normpath(self.test_images_path + '/X1_1990.jpg')
        self.x5_image_path = _path.normpath(self.test_images_path + '/X5_2007_2.jpg')

        img2pdf_layout = img2pdf.get_layout_fun((img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297)))





    #@unittest.skip("Temporaily disabled while debugging")
    def test_merge_image(self):
        """test merge"""
        rootdir = 'C:/temp/imgs'
        pdf, tmpdir, files = topdf.merge_img(rootdir, label_with_file=True, keep_tmp_images=True, overwrite=True)
        

    #@unittest.skip("Temporaily disabled while debugging")
    def test_merge_pdfs(self):
        """test merge pdfs"""
        rootdir = 'C:/temp/pdfs'
        pdfname = topdf.merge_pdf(rootdir, recurse=False, outfile='C:/temp/pdfs/all.pdf')
        pass


if __name__ == '__main__':
    """main"""
    unittest.main(verbosity=2)
