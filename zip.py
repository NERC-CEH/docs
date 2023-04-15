# pylint: disable=C0103, too-few-public-methods, locally-disabled, no-self-use, unused-argument
"""
Work with file archives using the zipfile package.

Methods:
    extract: Extract files from an archive (see method documentation)
    zipdir: Zip a directory, not recursive (see method documentation)
"""

import os.path as _path
import zipfile as _zipfile

import funclite.iolib as _iolib
import funclite.stringslib as _stringslib


class ExceptionNotAZip(Exception):
    """The file was not recognised as a zip file"""


def extract(fname: str, saveto: str,
            match_folder_name: (str, list, tuple, None) = None,
            match_file_name: (str, list, tuple, None) = None,
            match_file_ext: (str, list, tuple, None) = None):
    """
    Extract all files

    Args:
        fname (str): zip file name
        saveto (str): root folder to save extracted files to.
        match_folder_name (str, list, tuple, None): string or iterable containing simple strings to match on to the folder name (excluding any part of the file name)
        match_file_name (str, list, tuple, None): string or iterable containing simple strings to match on to the file name (excluding the dotted extension)
        match_file_ext (str, list, tuple, None): iterable of dotted file extensions to match

    Returns: None

    Notes: 
        This will silently overwrite files, so advise using a temporary folder to extract to.
        docs.filetypes has iterables of common file extensions.
        Any exception will stop extraction

    Examples:
        >>> zip.extract('./myzip.zip', match_folder_name=('folder',), match_file_name=('1',), match_file_ext=('.doc',))  # noqa
        >>> zip.extract('./myzip.zip', match_file_name=('1',), match_file_ext=filetypes.Images.dotted)  # noqa
    """
    saveto, fname = tuple(_path.normpath(p) for p in (saveto, fname))
    _, zipname, _ = _iolib.get_file_parts(fname)

    if not _zipfile.is_zipfile(fname):
        raise ExceptionNotAZip('{} is not a valid zipfile'.format(fname))

    with _zipfile.ZipFile(fname, mode='r') as Z:  # open the zipfile
        for ZI in Z.infolist():  # iterate over ZipInfo objects in zipfile
            if ZI.is_dir():
                continue

            fld, fname, ext = _iolib.get_file_parts(ZI.filename)

            if isinstance(match_file_name, str): match_file_name = [match_file_name]
            if match_file_name and not _stringslib.iter_member_in_str(fname, match_file_name, ignore_case=True):
                continue

            if isinstance(match_folder_name, str): match_folder_name = [match_folder_name]
            if match_folder_name and not _stringslib.iter_member_in_str(fld, match_folder_name, ignore_case=True):
                continue

            if isinstance(match_file_ext, str): match_file_ext = [match_file_ext]
            if match_file_ext and ext not in match_file_ext:
                continue

            Z.extract(ZI, path=saveto)


def zipdir(root: str, dest: str, name_match: (list, tuple, str, None) = (), overwrite: bool = True):
    """
    Zip a directory, does not recurse

    Args:
        root (str): Directory to zip
        dest (str): Zip to create
        name_match (str, list, None, tuple): File name matches
        overwrite (bool): overwrite

    Raises:
         FileExistsError: if not overwrite and dest exists

    Returns: None

    Notes:
        normpaths root and dest before using them
    """
    def _filt(t: str):
        return _stringslib.iter_member_in_str(t, name_match)

    dest = _path.normpath(dest)
    root = _path.normpath(root)
    if isinstance(name_match, str): name_match = [name_match]

    if _iolib.file_exists(dest) and not overwrite:
        raise FileExistsError('file %s exists. Use -o to overwrite.' % dest)

    fnames = list(filter(_filt, _iolib.file_list_generator1(root, '*', recurse=False)))

    PP = _iolib.PrintProgress(iter_=fnames)
    with _zipfile.ZipFile(dest, 'w', _zipfile.ZIP_DEFLATED) as Z:
        for f in fnames:
            d, fn, _ = _iolib.get_file_parts2(f)
            s = '%s/%s' % (_path.basename(d), fn)
            Z.write(f, s)
            PP.increment()
