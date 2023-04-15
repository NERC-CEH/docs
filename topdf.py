# pylint: disable=C0103, too-few-public-methods, locally-disabled, no-self-use, unused-argument
"""Routines for the manipulation of various documents"""

import os.path as _path
import shutil as _shutil

from PyPDF2 import PdfFileMerger
import numpy as _np
import cv2 as _cv2
import img2pdf as _img2pdf

import funclite.iolib as _iolib
import funclite.baselib as _baselib

import opencvlite.transforms as _transforms
import opencvlite.common as _common
import opencvlite.info as _info
from opencvlite.imgpipes import generators as _gen

_cInfo = _info.ImageInfo()


class a4:
    # x,y
    dpi72 = (595, 842)
    dpi150 = (1240, 1754)
    dpi300 = (2480, 3508)
    dpi600 = (4960, 7016)
    size_cm = (21.0, 29.7)
    size_inches = (8.3, 11.7)
    layout = _img2pdf.get_layout_fun((_img2pdf.mm_to_pt(210), _img2pdf.mm_to_pt(297)))
    layout_margin = _img2pdf.get_layout_fun((_img2pdf.mm_to_pt(168), _img2pdf.mm_to_pt(238)))


def _rotate_img(im):
    """(ndarray) -> ndarry
    helper routine, rotate image so it fits on portrait a4
    im:ndarry (the image)

    returns: ndarray (the rotated image)
    """
    w, h = _cInfo.resolution(im)
    cim = _np.copy(im)
    if w > h:
        cim = _transforms.rotate(cim, -90)
    return cim


def _build_lbl(s: str):
    """placeholder to build label for img"""
    return s


def merge_pdf(rootdir: str, find: (str, tuple, list, None) = None, exclude: (str, tuple, list, None) = None, out_file: str = '', sort_func: any = lambda s: s, recurse: bool = False) -> str:
    """
    Finding and merge all pdfs in rootdir

    Args:
        rootdir (str): root folder to find pdfs
        find (list, str, tuple, None): list or string, to match file names
        exclude (list, str, tuple, None): list or string, to exclude file names. Overrides find.
        out_file (str, None): full qualified name to call the merged pdf. Will be created in root dir with the basename of rootdir if out_file is empty
        sort_func (any): A function used to sort the order
        recurse (bool): recurse subdirs of rootdir

    Returns:
        str: filename of merged pdf

    Examples:
        Merge all pdfs in root of C:/temp, ordering by the file name length
        >>> merge_pdf('C:/temp', sort_func=len)
        'C:/temp/temp.pdf'
    """
    rootdir = _path.normpath(rootdir)
    basefld = _path.basename(rootdir)
    pdfs = []
    if isinstance(find, str): find = [find]

    for _, f, _, fname in _iolib.file_list_generator_dfe(rootdir, '*.pdf', recurse=recurse):

        # found is first used as a flag to determine if we have found text IN exclude
        found = False
        if exclude:
            for _ in exclude:
                found = _baselib.list_member_in_str(f, exclude)  # i.e. have we found anything that wildcard matches text IN exclude
                if found:
                    continue
        if found: continue

        found = True
        if find:
            found = False
            for _ in find:
                found = _baselib.list_member_in_str(f, find)
                if found: continue
        if not found: continue
        pdfs.append(fname)

    pdfs.sort(key=sort_func)

    merger = PdfFileMerger()
    _ = [merger.append(open(pdf, 'rb')) for pdf in pdfs]

    if out_file == '':
        out_file = _path.normpath('%s/%s%s' % (rootdir, basefld, '.pdf'))

    with open(out_file, 'wb') as fout:
        merger.write(fout)
    return out_file


def merge_pdf_by_list(pdfs: (list, tuple), out_file: str) -> None:
    """(iter, str, str) -> void
    Finding and merge all pdfs in the iterable pdfs

    Args:
        pdfs (list, tuple): iterable of fully qualified pdf file names
        out_file (str): full qualified name to call the merged pdf

    Returns: None

    Examples:
        >>> pdfs_ = ('c:\temp\1.pdf', 'c:\temp\2.pdf', 'c:\temp\3.pdf')
        >>> merge_pdf_by_list(pdfs_, 'c:\temp\1_2_3.pdf')
    """
    merger = PdfFileMerger()
    _ = [merger.append(open(_path.normpath(pdf), 'rb')) for pdf in pdfs]

    with open(_path.normpath(out_file), 'wb') as fout:
        #  print("If exporting to a network drive it may take a few minutes for the file to complete writing ... give it some time!")
        merger.write(fout)


def merge_img(rootdir: str, find: str = '', save_to_folder: str = '', pdf_file_name: str = '', overwrite: bool = False,
              label_with_file: bool = False, label_with_fld: bool = False, keep_tmp_images: bool = False,
              sorted_key: any = len) -> tuple[(str, None), (str, None), (list[str], None)]:
    """
    Find all images in rootdir then merge into a single pdf.
    This will correctly order the images according to the sorted_key func (len by default)

    Args:
        rootdir: root folder
        find: wildcard match for files
        save_to_folder (str): By default, saves the pdf to rootdir, use this to specify a different save folder
        pdf_file_name: save the pdf to this name. If ommitted will use the base folder name of rootdir
        overwrite: If the pdf output file name already exists, overwrite it
        label_with_file: label each image with the images filename
        label_with_fld: add the folder base name to the label
        keep_tmp_images: do not delete the temporary images folder
        sorted_key: function passed to key argument of the sorted function, used to set the file order.

    Returns:
        tuple[(str, None), (str, None), (list[str], None)]: pdf file name, temporary folder used to save adjusted images, list of matched images.
        Returns None,None,None if no images were found.

    Examples:
        Merge images in C:/temp/images saving as C:/mydir/merged.pdf and overriding the default "len" sort with a lambda expression
        >>> _, _, _ = merge_img('C:/temp/images', find='holiday', save_to_folder='C:/mydir', pdf_file_name='merged.pdf', sorted_key=lambda s:s)  # noqa
        ('C:/mydir/merged.pdf', ...
    """
    rootdir = _path.normpath(rootdir)
    if sorted_key is None: sorted_key = lambda x: x
    files = list(_iolib.file_list_generator1(rootdir, _common.IMAGE_EXTENSIONS_WILDCARDED))
    files.sort(key=sorted_key)
    if not files:
        print('No images in folder %s.' % rootdir)
        return None, None, None  # noqa

    base_fld = _path.basename(rootdir)
    tmp_fld = _iolib.temp_folder()  # get temp folder now so we can use it later
    _iolib.create_folder(tmp_fld)
    FL = _gen.FromList(files)

    PP = _iolib.PrintProgress(iter_=files)

    for img, pth, _ in FL.generate():
        if img is None or not img:
            break

        img = _rotate_img(img)

        w, h = a4.dpi72
        img = _transforms.resize(img, height=h, do_not_grow=True)
        img = _transforms.resize(img, width=w, do_not_grow=True)

        _, fname, _ = _iolib.get_file_parts(pth)

        if find:
            if not find.lower() in fname.lower():
                PP.increment()
                continue

        # img = _transforms.histeq_color(img)

        lbl = []
        if label_with_fld: lbl.append(base_fld)
        if label_with_file: lbl.append(_build_lbl(fname))

        if lbl:
            s = ' '.join(lbl)
            _common.draw_str(img, 10, 10, s, color=(0, 0, 0), scale=1.2, box_background=(255, 255, 255))

        fname = _path.normpath(tmp_fld + '/' + _iolib.get_temp_fname(suffix='.png', name_only=True))  # png expected to give better results
        # write the image to a temporary folder
        if not (_cv2.imwrite(fname, img)):  # noqa
            raise Exception('OpenCV failed to write %s to disk' % fname)
        files.append(fname)
        PP.increment()

    # now make the pdf
    if not pdf_file_name:
        if save_to_folder:
            pdf_file_name = '%s/%s' % (save_to_folder, base_fld + '.pdf')
        else:
            pdf_file_name = '%s/%s' % (rootdir, base_fld + '.pdf')

    pdf_file_name = _path.normpath(pdf_file_name)

    if not overwrite and _iolib.file_exists(pdf_file_name):
        raise FileExistsError('PDF file %s exists.' % pdf_file_name)

    print('Now creating the pdf, this may take sometime...')
    with open(pdf_file_name, "wb") as f:
        f.write(_img2pdf.convert(files, layout_fun=a4.layout_margin))

    if not keep_tmp_images:
        try:
            _shutil.rmtree(tmp_fld, ignore_errors=True)
        except:
            pass

    return pdf_file_name, tmp_fld, files

# This next routine looks complicatd, in essence it creates ordered lists of image file names, and then uses those lists to export the images, creating a matching ordered lists of the temporaryily created image files that is then merged into a single pdf document
def merge_img2(rootdir: str,
               find: (str, tuple, list, None) = None, exclude: (str, tuple, list, None) = None,
               save_to_folder: str = '',
               pdf_file_name: str = '',
               overwrite: bool = False,
               label_with_file: bool = False, label_with_fld: bool = False,
               keep_tmp_images: bool = False,
               file_sort_func: any = len,
               dir_sort_func: any = len,
               recurse: bool = False) -> tuple[(str, None), (str, None), (list[str], None)]:
    """
    Find all images in rootdir then merge into a single pdf.

    Differs from merge_img as it supports find and exclude as wildcard lists.

    Args:
        rootdir (str): root folder
        find (list, str, tuple, None): list or string, to match file names
        exclude (list, str, tuple, None): list or string, to exclude file names. Overrides find.
        save_to_folder (str): By default, saves the pdf to rootdir, use this to specify a different save folder
        pdf_file_name (str): save the pdf to this name. If ommitted will use the base folder name of rootdir
        overwrite (bool): If the pdf output file name already exists, overwrite it
        label_with_file (bool): label each image with the images filename
        label_with_fld (bool): add the **base** folder name to the label. Useful when it is the folders that have a pertinent name
        keep_tmp_images (bool): do not delete the temporary images folder
        file_sort_func (any): function passed to key argument of the sorted function, used to set the file order.
        dir_sort_func (any): function passed to key argument of the sorted function, applied to folders when using recurse.
        recurse (bool): recurse rootdir

    Returns:
        tuple[(str, None), (str, None), (list[str], None)]: The pdf file name, the temporary folder used to save adjusted images to and a list of matched images.
        Returns None,None,None if no images were found.

    Notes:
        Efforts are made to place in sensible order. But this will probably throw up the occasional misorder
        when directories are named inconsistently, e.g. 10, 009
        When usin

    Examples:
        Merge images in C:/temp/images saving as C:/mydir/merged.pdf and overriding the default "len" sort with a lambda expression
        >>> _, _, _ = merge_img('C:/temp/images', find='holiday', save_to_folder='C:/mydir', pdf_file_name='merged.pdf', sorted_key=lambda s:s)  # noqa
    """
    rootdir = _path.normpath(rootdir)
    base_fld = _path.basename(rootdir)
    if file_sort_func is None: file_sort_func = lambda x: x
    if dir_sort_func is None: dir_sort_func = lambda x: x

    files = []  # noqa

    #  Code for file ordering
    if recurse:
        folders = list(_iolib.folder_generator2(rootdir))
        del folders[folders.index(rootdir)]
        folders.sort(key=dir_sort_func)
        for fld in folders:
            lst = list(_iolib.file_list_generator2(fld, _common.IMAGE_EXTENSIONS_WILDCARDED, find=find, exclude=exclude, recurse=True))
            if lst:
                lst.sort(key=file_sort_func)
                files += lst  # Dont recurse deeper than 1 level)
    else:
        files = list(_iolib.file_list_generator2(rootdir, _common.IMAGE_EXTENSIONS_WILDCARDED, find=find, exclude=exclude, recurse=False))
        files.sort(key=file_sort_func)


    if not files:
        print('No images in folder %s.' % rootdir)
        return None, None, None  # noqa

    tmp_fld = _iolib.temp_folder()  # get temp folder now so we can use it later
    _iolib.create_folder(tmp_fld)
    FL = _gen.FromList(files)

    PP = _iolib.PrintProgress(iter_=files)
    image_files = []
    for img, pth, _ in FL.generate():
        img = _rotate_img(img)

        w, h = a4.dpi72
        img = _transforms.resize(img, height=h, do_not_grow=True)
        img = _transforms.resize(img, width=w, do_not_grow=True)

        lbl = []
        imgfld, fname, _ = _iolib.get_file_parts(pth)
        if label_with_fld: lbl.append(_path.basename(imgfld))
        if label_with_file: lbl.append(_build_lbl(fname))

        if lbl:
            s = ' '.join(lbl)
            _common.draw_str(img, 10, 10, s, color=(0, 0, 0), scale=1.2, box_background=(255, 255, 255))

        fname = _path.normpath(tmp_fld + '/' + _iolib.get_temp_fname(suffix='.png', name_only=True))  # png expected to give better results

        # write the image to a temporary folder
        if not (_cv2.imwrite(fname, img)):  # noqa
            raise Exception('OpenCV failed to write %s to disk' % fname)

        image_files.append(fname)  # append name so that we retain the correct order
        PP.increment()

    # now make the pdf
    if not pdf_file_name:
        if save_to_folder:
            pdf_file_name = '%s/%s' % (save_to_folder, base_fld + '.pdf')
        else:
            pdf_file_name = '%s/%s' % (rootdir, base_fld + '.pdf')

    pdf_file_name = _path.normpath(pdf_file_name)

    if not overwrite and _iolib.file_exists(pdf_file_name):
        raise FileExistsError('PDF file %s exists.' % pdf_file_name)

    print('Now creating the pdf, this may take sometime...')
    with open(pdf_file_name, "wb") as f:
        f.write(_img2pdf.convert(image_files, layout_fun=a4.layout_margin))

    if not keep_tmp_images:
        try:
            _shutil.rmtree(tmp_fld, ignore_errors=True)
        except:
            pass

    return pdf_file_name, tmp_fld, files


if __name__ == '__main__':
    pass
    # merge_pdf(r'C:\TEMP\bto\aerial_a3', out_file=r'C:\TEMP\bto\aerial_a3.pdf')
    # merge_pdf(r'C:\TEMP\bto\crn_2022_03_22', out_file=r'C:\TEMP\bto\crn_2022_03_22.pdf')
    # merge_pdf(r'C:\TEMP\bto\os_a3_5k', out_file=r'C:\TEMP\bto\os_a3_5k.pdf')
    # merge_pdf(r'C:\TEMP\bto\traffic_light_2022_03_22', out_file=r'C:\TEMP\bto\traffic_light_2022_03_22.pdf')
