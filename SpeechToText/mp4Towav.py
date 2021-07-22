import moviepy.editor as mp

mp4_path = r"C:\Users\BRPL\Videos\Captures\Cisco Webex Meetings 2021-06-04 16-01-31.mp4"
destination_path = r'D:\Process Improvement Project\python_programming\SpeechToText\wav_file\{}'


my_clip = mp.VideoFileClip(mp4_path)

my_clip.audio.write_audiofile(destination_path.format('result2.wav'))