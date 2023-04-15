"""Work with video files"""
import os.path as _path
import moviepy.editor as _editor




def concatenate(video_clip_paths: (list, tuple), output_path: str, method: str = "compose", **kwargs):
    """
    Concatenates several video files into one video file
    and save it to `output_path`.
    Note that extension (mp4, etc.) must be added to `output_path`

    Args:
        video_clip_paths: iterable of clips to concat
        output_path: path to save concate movie to

        method (str):
            reduce: Reduce the quality of the video to the lowest quality on the list of video_clip_paths.
            compose: type help(concatenate_videoclips) for the info
        kwargs (any): Keyword args, passed to moviepy.editor.concatenate_videoclips. See moviepy documentation.

    Returns: None
    """

    clips = [_editor.VideoFileClip(_path.normpath(c)) for c in video_clip_paths]
    if method == "reduce":
        # calculate minimum width & height across all clips
        min_height = min([c.h for c in clips])
        min_width = min([c.w for c in clips])
        # resize the videos to the minimum
        clips = [c.resize(newsize=(min_width, min_height)) for c in clips]
        # concatenate the final video
        final_clip = _editor.concatenate_videoclips(clips, **kwargs)
    elif method == "compose":
        # concatenate the final video with the compose method provided by moviepy
        final_clip = _editor.concatenate_videoclips(clips, method="compose", **kwargs)
    final_clip.write_videofile(output_path)  # noqa



if __name__ == '__main__':
    fnames = [r'C:\Users\gramon\Videos\2022-08-19 09-01-02.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-02-56.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-03-50.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-08-43.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-09-44.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-10-34.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 08-59-48.mp4',
              r'C:\Users\gramon\Videos\2022-08-19 09-00-16.mp4']
    concatenate(fnames, 'C:/TEMP/surveyors_how_to_access_sharepoint.mp4')
