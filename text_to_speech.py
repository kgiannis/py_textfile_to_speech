from win32com.client import Dispatch
import sys

""" Install necessary package: 
    pip install pypiwin32
 """

# Read file from command line
file_path = sys.argv[1]

speaker = Dispatch("SAPI.SpVoice")

# Read each line in file and convert it to speech
with open(file_path, "r", 500, "UTF-8") as fo:
    try:
        # Insert each line in array and calculate array length
        file_lines_in_array = fo.readlines()
        total_lines = len(file_lines_in_array)

        print("Total Lines: %d" % total_lines)
        print("---== Iterate Lines and convert to Speech ==---")

        # Read each line in array and convert to speech
        for item in file_lines_in_array:
            print(item)
            speaker.Speak(item)

        print("---== Transcript Completed ==---")

    except UnicodeDecodeError as err:
        print(err.object)

del speaker


