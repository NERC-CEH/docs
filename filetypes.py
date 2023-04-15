# pylint: disable=C0103, too-few-public-methods, locally-disabled, no-self-use, unused-argument
"""This module contains listing of file types by extension"""


class _BaseFileType():
    _extensions = ()

    @classmethod
    def extensions(cls):
        """bare extensions tuple
        """
        return cls._extensions

    @classmethod
    def dotted(cls):
        """() -> tuple
        Returns dotted tuple of file extensions, e.g. ('.bmp','jpeg',...)
        """
        return tuple('.{}'.format(s) for s in cls._extensions)

    @classmethod
    def asterixed(cls):
        """() -> tuple
        Returns dotted tuple of file extensions, e.g. ('*.bmp','*.jpeg',...)
        """
        return tuple('*.{}'.format(s) for s in cls._extensions)


class Images(_BaseFileType):
    """Raster file types"""
    _extensions = ('bmp', 'jpg', 'jpeg',
                   'png', 'tif', 'tiff', 'pbm', 'pgm', 'ppm', 'gif')


class Zip(_BaseFileType):
    """pkzip compatible"""
    _extensions = ('zip', 'pkzip', 'rar', 'arj', 'lzh')  # specific compatibility for any single file cannot be known


class MSWord(_BaseFileType):
    _extensions = ('doc', 'docx', 'docm', 'dot', 'dotm', 'dotx', 'odt', 'rtf')


class Text(_BaseFileType):
    _extensions = ('txt', 'csv')


class MSExcl(_BaseFileType):
    _extensions = ('xls', 'xlsm', 'csv', 'xlsb', 'xlsx', 'xlt', 'xltm', 'xltx', 'xlw')
