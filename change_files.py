import os, getopt, sys


opts, args = getopt.getopt(sys.argv[1:], "o:d:f:h")

offset, data, path = None, None, None
for name, value in opts:
    if name == "-o":
        offset = value.split(":")[1]
    elif name == "-d":
        data = value.split(":")[1]
    elif name == "-f":
        path = value[1:]
    elif name == "-h":
        print("\nThis script changes files in the specified path to a folder.\n"
              "Along with the path (-f), an offset (-o) and amount of data (-d) in bytes must\nbe specified.\n"
              '\nUsage:>> scriptname -o:523264 -d:1024 -f:"C:/files/"\n'
              'Excplanation:\n'
              'E.g. you have a folder with some 512 KB (524288 B) files and you need to change the last 1024 bytes\n'
              'of each file. So, offset in this case = 523264 and data to write = 1024. \n'
              'in this case, only the tail which is equal to 1KB of each file will be changed.\n'
              'If to specify offset bigger than the size of existing files, those files will be re-written to\n'
              'the specified size.\nEND.\n')


def file_changer(folder_path, offset_val, data_amount):

    global list_of_files_in_specified_folder
    global count_files_in_total  # Define global var to return it and use in further code
    list_of_files_in_specified_folder = []
    count_files_in_total = 0

    for (dirpath, dirnames, filenames) in os.walk(folder_path):
        list_of_files_in_specified_folder.extend(filenames)
        break

    for i in list_of_files_in_specified_folder:
        f = open(folder_path + i, 'rb+')
        f.read()
        f.seek(offset_val)
        f.write(os.urandom(data_amount))
        f.close()
        print("File '" + i + "' was modified")
        count_files_in_total += 1
    print("\n=============\n"+str(len(list_of_files_in_specified_folder)) + " file(s) processed")


if __name__ == '__main__':
    try:
        file_changer(path, int(offset), int(data))
    except TypeError:
        print("Some arguments were missed. Command must have 3 arguments! For help, use -h")
    except OSError:
        print("Incorrect value for argument")
    except ValueError:
        print("Value for argument is not specified")
    except OverflowError:
        print("Too big value of data specified")
    except KeyboardInterrupt:
        print("\n=============\nOperation cancelled")
        print(count_files_in_total, " files of " + str(len(list_of_files_in_specified_folder)) + " processed in total")
