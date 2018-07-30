# How to use:
# 1. Generate some files (0.5 - 1MB size) in a folder, e.g. C:\files
# 2. In this script, change the folder paths if the path with files is different;
# 3. Specify value of "file_offset" variable to set a point from where to start changing each file in the folder
# 4. In the function "os.urandom", Specify how many bytes you need to write to each file

import os
import sys


print("\nStat of the folder specified: " + str(os.stat("C:\ProgramData\TEMP")))
print("FS encoding: " + sys.getfilesystemencoding())
print("Username: " + os.getlogin() + "\n==============")
print("Platform: " + sys.platform.upper() + " " + os.name.upper())
print(os.getcwd() + " - is a current working directory and it contents: ")
print(os.listdir())
os.chdir("C:\\Programdata\\Temp\\")  # change current directory
print(os.getcwd() + " - Now it is a current working directory and it contents: ")
print(os.listdir())
trees = os.walk(os.getcwd())
print(trees)
print("WALK function output:")
for i in trees:
    print(i)
if os.path.exists("AAA"):
    try:
        os.remove("AAA")  # like "os.removedirs" which
    except PermissionError:
        print("It is a FOLDER, not a FILE!! Use 'rmdir()' to delete folders or 'os.removedirs()' to delete recursively.")
else:
    os.makedirs("AAA")  # like mkdir(), but makes all intermediate-level directories
print("os.truncate(path, lenght)" +" --> "+" Cuts a file to the lenght specified")
print(os.system("hostname"))



def file_changer(folder_path, offset_val, data_amount):


    file_list = []

    for (dirpath, dirnames, filenames) in os.walk(folder_path):
        file_list.extend(filenames)
        break

    for i in file_list:
        f = open(folder_path + i, 'rb+')
        f.read()
        print("File '" + i + "' was modified")
        f.seek(offset_val)
        f.write(os.urandom(data_amount))
        f.close()
    print("\n=============\n"+str(len(file_list)) + " file(s) processed")


if __name__ == '__main__':
    pass
    #file_changer("C:\\files\\", 1038576, 10000)  # path, offset, data
