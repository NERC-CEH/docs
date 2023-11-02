"""Work with file metadata"""
# Extent to include windows tags at some point - candidate library https://github.com/james-see/iptcinfo3
import os.path as _path
from dateutil.parser import parse as _parse

from PIL import Image as _Image
from PIL import ExifTags as _ExifTags






class Image:
    """
    Load exif data for an image.

    Args:
        fname: Image path (normpathed)

        filter_keys:
            A function used to filter the returned keys
            For example, filter_keys=lambda f: f in ['Aperture', 'ShutterSpeed']

    Methods:
        date_taken:
            Get the date taken as a datetime instance

        year_taken:
            Get the year taken

    """
    def __init__(self, fname: str, filter_keys=lambda f: True):
        self._fname = _path.normpath(fname)
        self.tags = {}
        self.gps_tags = {}
        self._filter_keys = filter_keys
        self._image_tags()
        self._date_taken = None


    def _image_tags(self):
        """
        reads all exif tags into tags
        """
        img = _Image.open(self._fname)
        self.tags = {_ExifTags.TAGS[k]: v for k, v in img._getexif().items() if k in _ExifTags.TAGS and self._filter_keys(_ExifTags.TAGS[k])}  # noqa

    @property
    def date_taken(self) -> "datetime.datetime":  # noqa
        """Date taken as a datetime instance

        Returns none if tags couldnt be read
        """
        if not self.tags: return
        if self._date_taken: return self._date_taken
        for dt in (self.tags['DateTime'], self.tags['DateTimeOriginal'], self.tags['DateTimeDigitized']):
            try:
                self. _date_taken = _parse(dt.replace(':', '').replace('-', ''))  # parse makes a good job of guessing, but colons across the board as a sep nakes it fail
                break
            except:
                pass

    @property
    def year_taken(self) -> int:
        """Return year taken"""
        return self._date_taken.year





if __name__ == '__main__':
    Img = Image(r'\\nerctbctdb\shared\shared\GMEP_Restricted\WP3_Completed squares\Year4\MainSurvey\Photos\28255_VegPlots_14_D_5.jpg')
    tags = Img.tags
    pass