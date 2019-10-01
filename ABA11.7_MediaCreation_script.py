"""
Script is created to automate Acronis Linux-based Bootable media AB11.7 creation (ISO) using Acronis Media Builder (MB).
The script uses GUI of the MB, so the mouse and the keyboard will be locked while creating ISOs.
The script can automatically install MB -> create ISOs -> uninstall MB of existing localization.
The script detects if there are MSI installers of the MB in "installers" folder. If there aren't, all MSI files of MB
will be unpacked into that folder from the detected BIG installers. After that, the script must be restarted.
How To Use it:
    1. Prepare a VM Win7x64 with 1 CD-ROM and 1 network (DHCP), w/o Floppy or USB flash! Disk C: = 100GB
    2. Make sure that na_keys.json and keys.json are present in the root folder near the script. They have licenses.
    3. On the machine, run Setup.bat to setup Python 3 with pywinauto lib and 7zip, and correct system settings
    4. Put all big installers of ABR to /installers folder and run Unzip.bat. MSI will be extracted to separate folders
    5. Run the script:> python ABA11.7_MediaCreation_script.py or by double-clicking
    6. Wait when all ISOs of all localizations are created (find them in '/media/' folder).
    7. Create ISOs with TRIAL keys manually (actual for US localization only)
"""
from pywinauto.application import Application
from pywinauto import application
from pywinauto import base_wrapper
from pywinauto import MatchError
from pywinauto import ElementNotFoundError
from pywinauto import timings
import os
import time
import logging
import json
import sys
from ctypes import *

start_time = time.time()
file_formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
console_formatter = logging.Formatter('%(asctime)s %(message)s')


def setup_logger(name, log_file, level=logging.INFO):  # Function setup as many loggers as you want
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(file_formatter)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.ERROR)
    console_handler.setLevel(logging.WARNING)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


# Creating a log file
debuglog = setup_logger('debuglog', 'C:\MediaCreationLog.log')
debuglog.info("Logging started...")


def get_build_and_edition(inst_folder_name):
    if inst_folder_name[26] == "N":
        return True, inst_folder_name[28:33]
    else:
        return False, inst_folder_name[27:32]


installer_folder_list = []  # Folders list
get_curr_dir = os.getcwd()

for (dirpath, dirnames, filenames) in os.walk(".\\installers"):  # Getting a folder list with installers
    installer_folder_list.extend(dirnames)
    break
if installer_folder_list == []:  # Installers existence check
    debuglog.warning('No MSI installers found')
    os.system(".\\Unzip.bat")
    debuglog.info('Check if the installers exist and RE-RUN THE SCRIPT')
else:
    debuglog.info("MSI installers exist: " + str(installer_folder_list) + " Continue...")

list_of_all_localization_symbols = []  # List that contains all symbols related to localization part of folder names
iteration = 0
list_of_grouped_localizations = []  # A list of lists of localizations
localization_list = []  # Ready localizations used for main script
local = ''
for i in installer_folder_list:  # Extending list_of_all_localization_symbols list, e.g. [en-US, fr-FR]
    if get_build_and_edition(installer_folder_list[0])[0]:
        a = i[34:]
    else:
        a = i[33:]
    list_of_all_localization_symbols.extend(a)
    while iteration < len(list_of_all_localization_symbols):  # Creating a list of grouped localizations list
        list_of_grouped_localizations.append(list_of_all_localization_symbols[iteration:iteration + 5])
        iteration += 5
        for new_cycle_to_get_local_list in range(
                len(list_of_grouped_localizations)):  # Finalizing localization list
            local = ''.join(list_of_grouped_localizations[new_cycle_to_get_local_list])
        localization_list.extend([local])
debuglog.info("Installers with the following localizations are found in the folder: " + str(localization_list))

try:
    if not get_build_and_edition(installer_folder_list[0])[0]:
        try:
            data = json.load(open(".\\keys.json", 'r'))  # Opening JSON with keys
            keys_list = data["main_keys"]
        except FileNotFoundError:
            debuglog.error("json file with keys not found! Put the file into the root folder along with the script.")
            exit()
    else:
        try:
            data = json.load(open(".\\na_keys.json", 'r'))
            na_keys_list = data["main_na_keys"]
        except FileNotFoundError:
            debuglog.error("json file with NA keys not found! Put the file into the root folder along with the script.")
            exit()
except IndexError:
    debuglog.error("No installers found in folder '\installers'")
    exit()


names_list = ["AcronisBackupAdvancedWS_11.7_", "AcronisBackupAdvancedUniversal_11.7_", "AcronisBackupAdvancedHyperV_11.7_", "AcronisBackupAdvancedVMware_11.7_", "AcronisBackupAdvancedRHEV_11.7_", "AcronisBackupAdvancedXEN_11.7_", "AcronisBackupAdvancedOracle_11.7_", "AcronisBackupEssentials_11.7_", "AcronisBackupAdvancedPC_11.7_", "AcronisBackupWS_11.7_", "AcronisBackupPC_11.7_"]
na_names_list = ['AcronisBackupAdvancedWS_11.7N_', 'AcronisBackupAdvancedUniversal_11.7N_', 'AcronisBackupAdvancedHyperV_11.7N_', 'AcronisBackupAdvancedVMware_11.7N_', 'AcronisBackupAdvancedRHEV_11.7N_', 'AcronisBackupAdvancedXEN_11.7N_', 'AcronisBackupAdvancedOracle_11.7N_', 'AcronisBackupEssentials_11.7N_', 'AcronisBackupAdvancedPC_11.7N_', 'AcronisBackupWS_11.7N_', 'AcronisBackupPC_11.7N_']
os.chdir(".\\media")  # Change current working directory to dir for ISOs
current_working_directory = os.getcwd()  # Getting the current work dir to use it in Installation func.

try:
    for i in range(len(localization_list)):  # Creating sub-folders in /media folder for ISO
        debuglog.info("Creating sub-folder " + localization_list[i] + " for the installer")
        os.makedirs(localization_list[i])
except FileExistsError:
    debuglog.info("Target sub-folders exist")


def uninstallation():
    """ Is called when uninstallation of Media Builder is required """
    debuglog.info("Media Builder uninstallation starts!")
    debuglog.info('Executing "wmic product where vendor="Acronis" call uninstall"')
    os.system('wmic product where vendor="Acronis" call uninstall')
    if not os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("Media Builder uninstallation complete!")
    else:
        debuglog.warning("Media Builder uninstallation FAILED! See MSI log for more information or try to uninstall the Media builder manually")


def installation(current_local):
    """ Is called when installation of Media Builder is required """
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe') or os.path.exists('C:\\Program Files\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("An unknown Media Builder is present in system and will be uninstalled.")
        uninstallation()
    debuglog.info("New Media Builder is being installed. Please wait...")
    parent_of_current_dir = os.path.abspath(os.path.join(current_working_directory, os.pardir))
    if get_build_and_edition(installer_folder_list[0])[0]:
        installation_command = 'start /wait msiexec /i ' + parent_of_current_dir + '\\installers\\AcronisBackupAdvanced_11.7N_' + get_build_and_edition(installer_folder_list[0])[1] + '_' + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    else:
        installation_command = 'start /wait msiexec /i ' + parent_of_current_dir + '\\installers\\AcronisBackupAdvanced_11.7_' + get_build_and_edition(installer_folder_list[0])[1] + '_' + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    try:
        debuglog.info("Executing " + installation_command)
        os.system(installation_command)
    except application.AppStartError:
        debuglog.error("Check if the installation command and the path to the installer are correct.")
        exit()
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("Installation complete!")
    else:
        debuglog.error("Installation failed! Check if 'C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe' exists")


def main_script(key_list_f, names_list_f, locale):
    """ Goes through the wizard and creates ISO files"""
    for k in range(len(key_list_f)):  # k is an index of a license. This is a loop of creating ISOs
        debuglog.warning("ISO creation starts...")
        new_iso_name = current_working_directory + "\\" + locale + "\\" + names_list_f[k] + get_build_and_edition(installer_folder_list[0])[1] + '_' + locale
        app = Application()
        try:
            app.start("C:\Program Files (x86)\Common Files\Acronis\MediaBuilder\MediaBuilder.exe")
            debuglog.info("Executing 'C:\Program Files (x86)\Common Files\Acronis\MediaBuilder\MediaBuilder.exe'")
        except application.AppStartError:
            debuglog.error("Check if the path to MediaBuilder.exe is correct.")
            exit()
        time.sleep(2)
        window = app.window()
        window.set_focus()
        buildwizard = app.BuilderWizard
        buildwizard.wait('ready')
        next_button = buildwizard[u'Next >FXButton3']
        debuglog.info("Clicking 'Next' button in the start Wizard page")
        next_button.click()
        next_button.wait('ready', timeout=20)
        debuglog.info("Clicking 'Next' button in the 'Select media type' page")
        next_button.click()
        next_button.wait('ready', timeout=20)
        debuglog.info("Clicking 'Next' button in the 'Kernel parameters' page")
        next_button.click()
        debuglog.info("Selecting AB x86/x64 components check boxes")
        buildwizard.FXAConfigurationTree1.click()
        buildwizard.send_keystrokes('{SPACE}')
        debuglog.info("Clicking 'Next' button in the 'Components' page")
        next_button.click()
        next_button.wait('ready', timeout=20)
        debuglog.info("Clicking 'Next' button in the 'Network and remote connection' page")
        next_button.click()
        debuglog.info("Clicking the first 'YES' button")
        buildwizard.send_keystrokes('{ENTER}')
        debuglog.info("Clicking the second 'YES' button")
        buildwizard.send_keystrokes('{ENTER}')
        debuglog.info("Selecting a radio button to enter a key")
        buildwizard.send_keystrokes('{DOWN}')
        buildwizard.send_keystrokes('{DOWN}')
        buildwizard.send_keystrokes('{SPACE}')
        debuglog.info("Selecting Text field for entering a key")
        fxtext = buildwizard['FXText']
        fxtext.click()
        fxtext.send_chars(key_list_f[k])
        fxtext.click()
        while True:  # Re-type key if it is not correct
            key_copy = fxtext.window_text()
            if key_copy.upper() != key_list_f[k].upper():
                debuglog.info("Entered key is " + key_copy)
                debuglog.warning("Entered key is NOT correct. Re-typing...")
                fxtext.click()
                fxtext.send_keystrokes('^a')
                fxtext.send_keystrokes('{DELETE}')
                fxtext.send_chars(key_list_f[k])
            else:
                debuglog.info("Entered key is " + key_copy)
                debuglog.info("Entered key is correct. Continue...")
                break
        time.sleep(0)
        window.set_focus()
        fxtext.wait('ready', timeout=20)
        debuglog.info("Clicking 'Next' button in the 'Enter a key' page")
        next_button.click()
        debuglog.info("Clicking 'Next' button in the 'Specified key' page")
        next_button.wait('ready', timeout=20)
        next_button.click()
        debuglog.info("Selecting a panel with a list of media targets")
        buildwizard.send_keystrokes('{TAB}')
        # ***SELECT "ISO" AS A MEDIA TYPE:*** (Here we select ISO item from a media type list according to localization)
        if locale == "en-US" or locale == "en-EU" or locale == "de-DE" or locale == "zh-TW" or locale == "ja-JP" or locale == "ko-KR" or locale == "pl-PL" or locale == "zh-CN" or locale == "es-ES":
            buildwizard.send_keystrokes('{UP}')
        elif locale == "cs-CZ" or locale == "it-IT":
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
        elif locale == "fr-FR":
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
        time.sleep(1)
        debuglog.info("Clicking 'Next' button in the 'Bootable Media format' page")
        next_button.set_focus()
        next_button.click()
        time.sleep(1)
        debuglog.info("Clicking a path field")
        buildwizard.FXAFileNameField.click()
        debuglog.info("Erasing the default ISO name in the field")
        buildwizard.FXAFileNameField.send_keystrokes('^a')
        buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
        time.sleep(0)
        debuglog.info("Specifying a new path and a file name")
        buildwizard.FXAFileNameField.send_chars(new_iso_name)
        while True:  # Re-type ISO name if it is not correct
            iso_name__copy = buildwizard.FXAFileNameField.window_text()
            if iso_name__copy.lower() != new_iso_name.lower():  # Checking if ISO name was correctly entered
                buildwizard.FXAFileNameField.click()
                buildwizard.FXAFileNameField.send_keystrokes('^a')
                buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
                time.sleep(0)
                buildwizard.FXAFileNameField.send_chars(new_iso_name)
                debuglog.warning("'" + new_iso_name + "' is NOT a correct ISO name. Re-typing")
            else:
                debuglog.info("'" + new_iso_name + "' is a correct ISO name. Continue...")
                break
        time.sleep(0)
        debuglog.info("Clicking 'Next' button in the 'Specify name for ISO' page")
        next_button.set_focus()
        next_button.wait('exists visible enabled ready active', timeout=20)
        next_button.set_focus()
        next_button.click()
        debuglog.info("Clicking 'Next' button in the 'Add drivers' page")
        next_button.click()
        debuglog.info("Clicking 'Proceed' button in the 'SUMMARY' page")
        next_button.wait('ready')
        next_button.set_focus()
        next_button.click()
        finish_message_box = app.FXAMessageBoxImpl
        finish_message_box.wait("ready active", timeout=50)
        ok_button_in_box = finish_message_box.FXButton
        debuglog.info("Clicking 'OK' button in the message box")
        ok_button_in_box.click()
        if os.path.exists(new_iso_name + ".iso"):  # ISO creation check
            debuglog.warning('Media is created \n====================')
        else:
            debuglog.error("ISO is not created!")


def start():  # Starts loop 'Uninstallation -> Installation -> Creating_ISO -> Uninstallation'
    for i in range(len(localization_list)):
        installation(localization_list[i])
        time.sleep(1)
        if get_build_and_edition(installer_folder_list[0])[0]:
            debuglog.info("'NORTH AMERICAN' build installed in the system")
            try:
                main_script(na_keys_list, na_names_list, localization_list[i])
            except (MatchError, ElementNotFoundError, base_wrapper.ElementNotEnabled, timings.TimeoutError):
                debuglog.error("Element Not Found or enabled! Timeout error. Re-run the script.")
                debuglog.info("--- Total duration  %s seconds ---" % round((time.time() - start_time)))
                exit()
        else:
            debuglog.info("'MAINTENANCE' build installed in the system")
            try:
                main_script(keys_list, names_list, localization_list[i])
            except (MatchError, ElementNotFoundError, base_wrapper.ElementNotEnabled, timings.TimeoutError):
                debuglog.error("Element Not Found or enabled! Timeout error. Re-run the script.")
                debuglog.info("--- Total duration %s seconds ---" % round((time.time() - start_time)))
                exit()
        time.sleep(1)
        uninstallation()


if __name__ == '__main__':
    debuglog.info("Locking mouse and keyboard")
    windll.user32.BlockInput(True)  # Block mouse/keyboard inputs
    start()
    debuglog.info("Operation completed! See the log file 'C:\MediaCreationLog.log' for more information.")
    debuglog.info("Create TRIAL US media manually!\n")

