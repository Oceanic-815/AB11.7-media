# How to use:
# 1. Generate some files (0.5 - 1MB size) in a folder, e.g. C:\files
# 2. In this script, change the folder paths if the path with files is different;
# 3. Specify value of "file_offset" variable to set a point from where to start changing each file in the folder
# 4. In the function "os.urandom", Specify how many bytes you need to write to each file

import os


file_list = []


for (dirpath, dirnames, filenames) in os.walk("C:\\files\\"):
    file_list.extend(filenames)
    break

for i in file_list:
    f = open("C:\\files\\" + i, 'rb+')
    contents = f.read()
    print("File " + i + " was modified")
    file_offset = f.seek(1038576)
    f.write(os.urandom(10000))
    f.close()
