"""
Script is created to automate Acronis Linux-based Bootable media AB11.7 creation (ISO) using Acronis Media Builder (MB).
The script uses GUI of the MB, so DO NOT use the mouse and the keyboard while creating ISOs. Better to use a separate VM
The script can automatically install MB -> create ISOs -> uninstall MB of existing localization.
The script detects if there are MSI installers of the MB in ".\installers" folder. If there aren't, all MSI files of MB
will be unpacked into that folder from the detected BIG installers.
How To Use it:
    1. Prepare a VM Win7x64 with 1 CD-ROM and 1 network (DHCP), w/o Floppy or USB flash! Disk C: = 100GB
    2. Specify new build number in the variable "build_number", e.g "50073"
    3. If NA build should be created, specify if_na = True, if Maint build - False
    4. Make sure that na_keys.json and keys.json are present in the root folder near the script. They have licenses.
    5. On the machine, run Setup.bat to setup Python 3 with pywinauto lib and 7zip, and correct system settings
    6. Put all big installers of ABR to /installers folder and run Unzip.bat. MSI will be extracted to separate folders
    7. Run the script:> python ABA11.7_MediaCreation_script.py
    8. Wait when all ISOs of all localizations are created (find them in '/media/' folder).
    9. Create ISOs with TRIAL keys manually (actual for US localization only)
"""

from pywinauto.application import Application
from pywinauto import application
from pywinauto import base_wrapper
from pywinauto import MatchError
from pywinauto import ElementNotFoundError
import os
import time
import logging
import json
import sys


build_number = "50420"  # Specify a build number with "_" character in the end. Example: "50064_"
if_na = True            # Set True if NA build is used


formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
console_formatter = logging.Formatter('%(asctime)s %(message)s')


def setup_logger(name, log_file, level=logging.INFO):  # Function setup as many loggers as you want
    handler = logging.FileHandler(log_file)
    handler.setFormatter(formatter)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    logger.addHandler(console_handler)
    return logger


# Creating a log file
debuglog = setup_logger('debuglog', 'C:\MediaCreationLog.log')
debuglog.info('Logging started')


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
    if if_na:
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

if not if_na:
    try:
        data = json.load(open(".\\keys.json", 'r'))  # Opening JSON with keys
        keys_list = data["main_keys"]
    except FileNotFoundError:
        debuglog.info("ERROR: json file with keys not found! Put the file into the root folder along with the script.")
        exit()
else:
    try:
        data = json.load(open(".\\na_keys.json", 'r'))
        na_keys_list = data["main_na_keys"]
    except FileNotFoundError:
        debuglog.info("ERROR: json file with NA keys not found! Put the file into the root folder along with the script.")
        exit()

names_list = ["AcronisBackupAdvancedWS_11.7_", "AcronisBackupAdvancedUniversal_11.7_", "AcronisBackupAdvancedHyperV_11.7_", "AcronisBackupAdvancedVMware_11.7_", "AcronisBackupAdvancedRHEV_11.7_", "AcronisBackupAdvancedXEN_11.7_", "AcronisBackupAdvancedOracle_11.7_", "AcronisBackupEssentials_11.7_", "AcronisBackupAdvancedPC_11.7_", "AcronisBackupWS_11.7_", "AcronisBackupPC_11.7_"]
na_names_list = ['AcronisBackupAdvancedWS_11.7N_', 'AcronisBackupAdvancedUniversal_11.7N_', 'AcronisBackupAdvancedHyperV_11.7N_', 'AcronisBackupAdvancedVMware_11.7N_', 'AcronisBackupAdvancedRHEV_11.7N_', 'AcronisBackupAdvancedXEN_11.7N_', 'AcronisBackupAdvancedOracle_11.7N_', 'AcronisBackupEssentials_11.7N_', 'AcronisBackupAdvancedPC_11.7N_', 'AcronisBackupWS_11.7N_', 'AcronisBackupPC_11.7N_']

os.chdir(".\\media")  # Change current working directory to dir for ISOs
current_working_directory = os.getcwd()  # Getting the current work dir to use it in Installation func.

try:
    for i in range(len(localization_list)):  # Creating subfolders in /media folder for ISO
        os.makedirs(localization_list[i])
except FileExistsError:
    debuglog.info("Target sub-folders exist")


def uninstallation():
    """ Is called when uninstallation of Media Builder is required """
    debuglog.info("Media Builder uninstallation starts!")
    os.system('wmic product where vendor="Acronis" call uninstall')
    if not os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("Media Builder uninstallation complete!")
    else:
        debuglog.warning("Media Builder uninstallation FAILED! See MSI log for more information or try uninstalling builder manually")


def installation(current_local):
    """ Is called when installation of Media Builder is required """
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("An unknown Media Builder is present in system and will be uninstalled.")
        uninstallation()
    debuglog.info("New Media Builder is being installed. Please wait...")
    parent_of_current_dir = os.path.abspath(os.path.join(current_working_directory, os.pardir))
    if if_na:
        installation_command = 'start /wait msiexec /i ' + parent_of_current_dir + '\\installers\\AcronisBackupAdvanced_11.7N_' + build_number + '_' + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    else:
        installation_command = 'start /wait msiexec /i ' + parent_of_current_dir + '\\installers\\AcronisBackupAdvanced_11.7_' + build_number + '_' + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    debuglog.info('Exexuting ' + installation_command)
    try:
        os.system(installation_command)
    except application.AppStartError:
        debuglog.info("Check if build number is correct. If NA build is used, specify 'if_na = True', else 'if_na = False'")
        exit()
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        debuglog.info("Installation complete!")
    else:
        debuglog.warning("Installation failed!")


def main_script(key_list_f, names_list_f, locale):
    """ Goes through the wizard and creates ISO files"""
    for k in range(len(key_list_f)):  # k is an index of a license. This is a loop of creating ISOs
        debuglog.info("ISO creation starts...")
        new_iso_name = current_working_directory + "\\" + locale + "\\" + names_list_f[k] + build_number + '_' + locale
        app = Application()
        try:
            app.start("C:\Program Files (x86)\Common Files\Acronis\MediaBuilder\MediaBuilder.exe")
        except application.AppStartError:
            debuglog.info("Check if build number is correct. If NA build installed, specify 'if_na = True', else 'if_na = False'")
            exit()
        time.sleep(2)
        window = app.window_()
        window.SetFocus()
        buildwizard = app.BuilderWizard
        buildwizard.Wait('ready')
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
            key_copy = fxtext.WindowText()
            if key_copy.upper() != key_list_f[k].upper():
                debuglog.warning("Entered key is NOT correct. Re-typing...")
                fxtext.click()
                fxtext.send_keystrokes('^a')
                fxtext.send_keystrokes('{DELETE}')
                fxtext.send_chars(key_list_f[k])
            else:
                debuglog.info("Entered key is correct. Continue...")
                break
        time.sleep(0)
        window.SetFocus()
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
        next_button.SetFocus()
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
            iso_name__copy = buildwizard.FXAFileNameField.WindowText()
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
        next_button.SetFocus()
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
            debuglog.info('Media is created \n====================')
        else:
            debuglog.warning("ISO is not created!")


def start():
    for i in range(len(localization_list)):
        installation(localization_list[i])
        time.sleep(1)
        if if_na:
            debuglog.info("'if_na = True' specified")
            try:
                main_script(na_keys_list, na_names_list, localization_list[i])
            except (MatchError, ElementNotFoundError, base_wrapper.ElementNotEnabled):
                debuglog.warning("Error: Element Not Found/Enabled! Please Re-run the script.")
                exit()
        else:
            debuglog.info("'if_na = False' specified")
            try:
                main_script(keys_list, names_list, localization_list[i])
            except (MatchError, ElementNotFoundError, base_wrapper.ElementNotEnabled):
                debuglog.warning("Error: Element Not Found/Enabled! Please Re-run the script.")
                exit()
        time.sleep(1)
        uninstallation()


if __name__ == '__main__':
    start()
    debuglog.info("Operation completed! See the log file 'C:\MediaCreationLog.log' for more information.")
    debuglog.info("Create TRIAL US media manually!\n")
