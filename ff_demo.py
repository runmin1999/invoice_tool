import ffmpeg


ffmpeg_path = "D:/ProgramDate/ffmpeg_full_build/bin/ffmpeg"
input_file = "E:/Code/ff_demo.mp4"
# output_file = "E:/Code/output.mp4"


# def segmentation_str(strs, length):
#     import re
#     return re.findall(".{"+str(length)+"}", strs)


def clip_video(input_file, output_file, ffmpeg_path, start, end):
    out_log, err_log = (
        ffmpeg
        .input(input_file, ss=start, to=end)
        .output(output_file, f="mp4", loglevel='error')
        .run(cmd=ffmpeg_path, capture_stdout=True, overwrite_output=True)
    )
    return out_log, err_log


def calculate_case_time(record, example_start, example_end):
    video_of_start = before_minutes_seconds(int(example_start)-int(record))
    video_of_end = before_minutes_seconds(int(example_end)-int(record))
    return video_of_start, video_of_end


def before_minutes_seconds(nums):
    return str(nums).rjust(6, '0')


def after_minutes_seconds(nums):
    return str(nums).ljust(6, '0')


if __name__ == "__main__":

    record = after_minutes_seconds("120000")

    i = 0

    for item in [[121010, 121030], [120510, 120512]]:

        i = i+1

        output_file = "E:/Code/" + str(i) + ".mp4"

        example_start = before_minutes_seconds(item[0])
        example_end = before_minutes_seconds(item[1])

        video_of_start, video_of_end = calculate_case_time(
            record, example_start, example_end)

        out_log, err_log = clip_video(
            input_file, output_file, ffmpeg_path, video_of_start, video_of_end)

        if out_log == b'':
            print('end')
