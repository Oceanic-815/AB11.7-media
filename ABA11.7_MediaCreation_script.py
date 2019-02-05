"""
Script is created to automate Acronis Linux-based Bootable media AB11.7 creation (ISO) using Acronis Media Builder (MB).
The script uses GUI of the MB, so DO NOT use the mouse and the keyboard while creating ISOs. Better to use a separate VM
The script can automatically install MB -> create ISOs -> uninstall MB of existing localization.
The script detects if there are MSI installers of the MB in ".\installers" folder. If there aren't, all MSI files of MB
will be unpacked into that folder from the detected BIG installers.
How To Use it:
    1. Prepare a VM with Win7 x64 with 1 CD-ROM and no Floppy or flash! It is important step! Disk C: = 100GB
    2. Specify new build number in the variable "build_number", e.g "50073"
    3. If NA build should be created, specify if_na = True, if Maint build - False
    4. Make sure that na_keys.json and keys.json are present in the root folder near the script. They have licenses.
    5. On the machine, run Setup.bat to setup Python 3 with pywinauto lib and 7zip, and correct system settings
    6. Put all big installers of ABR to ./installers folder and run Unzip.bat. MSI will be extracted to separate folders
    7. Run the script:> python ABA11.7_MediaCreation_script.py
    8. Wait when all ISOs of all localizations are created (find them in '../media/' folder).
    9. Create ISOs with TRIAL keys manually (actual for US localization only)
"""

from pywinauto.application import Application
import os
import time
import logging
import json

build_number = "50064"  # Specify a build number with "_" character in the end. Example: "50064_"
if_na = False            # Set True if NA build is used

try:
    logging.basicConfig(level=logging.INFO, filename='C:\MediaCreationLog.log', format='%(asctime)s %(message)s')
except:
    pass

installer_folder_list = []  # Folders list
get_curr_dir = os.getcwd()

for (dirpath, dirnames, filenames) in os.walk(".\\installers"):  # Getting a folder list with installers
    installer_folder_list.extend(dirnames)
    break
if installer_folder_list == []:  # Installers existence check
    print('No MSI installers found. Extracting MSI files...')
    logging.info('No MSI installers found. Extracting MSI files...')
    os.system(".\\Unzip.bat")
    print('\nMSI FILES EXTRACTED! RE-RUN THE SCRIPT\n')
    logging.info('MSI files extracted')
else:
    logging.info("MSI installers exist: " + str(installer_folder_list) + " Continue...")
    print("MSI installers exist. Continue...")

list_of_all_localization_symbols = []  # List that contains all symbols related to localization part of folder names
iteration = 0
list_of_grouped_localizations = []  # List of lists of localizations
localization_list = []  # Ready localizations used for main script
local = ''
for i in installer_folder_list:  # Extending list_of_all_localization_symbols list
    a = i[33:]
    list_of_all_localization_symbols.extend(a)
    while iteration < len(list_of_all_localization_symbols):  # Creating a list of grouped localizations list
        list_of_grouped_localizations.append(list_of_all_localization_symbols[iteration:iteration + 5])
        iteration += 5
        for new_cycle_to_get_local_list in range(
                len(list_of_grouped_localizations)):  # Finalizing localization list
            local = ''.join(list_of_grouped_localizations[new_cycle_to_get_local_list])
        localization_list.extend([local])
print(localization_list)
logging.info("Installers with the following localizations are found in the folder: " + str(localization_list))


try:
    data = json.load(open(".\\keys.json", 'r'))  # Opening JSON with licenses
    keys_list = data["main_keys"]
except FileNotFoundError:
    print("ERROR: json file with keys not found! Put the file into the root folder along with the script.")
    logging.info("ERROR: json file with keys not found!")

try:
    data = json.load(open(".\\na_keys.json", 'r'))
    na_keys_list = data["main_na_keys"]
except FileNotFoundError:
    print("ERROR: json file with NA keys not found! Put the file into the root folder along with the script.")
    logging.info("ERROR: File with NA keys not found!")

names_list = ["AcronisBackupAdvancedWS_11.7_", "AcronisBackupAdvancedUniversal_11.7_", "AcronisBackupAdvancedHyperV_11.7_", "AcronisBackupAdvancedVMware_11.7_", "AcronisBackupAdvancedRHEV_11.7_", "AcronisBackupAdvancedXEN_11.7_", "AcronisBackupAdvancedOracle_11.7_", "AcronisBackupEssentials_11.7_", "AcronisBackupAdvancedPC_11.7_", "AcronisBackupWS_11.7_", "AcronisBackupPC_11.7_"]
na_names_list = ['AcronisBackupAdvancedWS_11.7N_', 'AcronisBackupAdvancedUniversal_11.7N_', 'AcronisBackupAdvancedHyperV_11.7N_', 'AcronisBackupAdvancedVMware_11.7N_', 'AcronisBackupAdvancedRHEV_11.7N_', 'AcronisBackupAdvancedXEN_11.7N_', 'AcronisBackupAdvancedOracle_11.7N_', 'AcronisBackupEssentials_11.7N_', 'AcronisBackupAdvancedPC_11.7N_', 'AcronisBackupWS_11.7N_', 'AcronisBackupPC_11.7N_']

os.chdir(".\\media")
current_working_directory = os.getcwd()

try:
    for i in range(len(localization_list)):
        os.makedirs(localization_list[i])
except FileExistsError:
    print("Target folders exist")
    logging.info("Target folders exist")


def uninstallation():
    """ Is called when uninstallation of Media Builder is required """
    logging.info("Media Builder uninstallation starts!")
    print("Media Builder is being uninstalled. Please wait...")
    os.system('wmic product where vendor="Acronis" call uninstall')
    if not os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        logging.info("Media Builder uninstallation complete!")
        print("====================\nMedia Builder is uninstalled!")
    else:
        logging.info("Media Builder uninstallation FAILED!")
        print("Something went wrong! See MSI log for more information or try uninstalling builder manually")


def installation(current_local):
    """ Is called when installation of Media Builder is required """
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        logging.info("An unknown Media Builder is present in system and will be uninstalled.")
        uninstallation()
    print("New Media Builder is being installed. Please wait...")
    logging.info("New Media Builder is being installed.")
    parent_of_current_dir = os.path.abspath(os.path.join(current_working_directory, os.pardir))
    installation_command = 'start /wait msiexec /i '+parent_of_current_dir+'\\installers\\AcronisBackupAdvanced_11.7_' + build_number + '_' + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    print(installation_command)
    os.system(installation_command)
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        print("Installation complete!")
        logging.info("Installation complete!")
    else:
        print("Installation FAILED!")
        logging.info("Installation failed!")


def main_script(key_list_f, names_list_f, locale):
    """ This function creates ISO """
    for k in range(len(key_list_f)):  # k is an index of a license. This is a loop of creating ISOs
        logging.info("ISO creation starts...")
        new_iso_name = current_working_directory + "\\" + locale + "\\" + names_list_f[k] + build_number + '_' + locale
        app = Application().start("C:\Program Files (x86)\Common Files\Acronis\MediaBuilder\MediaBuilder.exe")
        time.sleep(2)
        window = app.window_()
        window.SetFocus()
        buildwizard = app.BuilderWizard
        buildwizard.Wait('ready')
        next_button = buildwizard[u'Next >FXButton3']
        next_button.click()
        logging.info("'Next' button in the start Wizard page was clicked")
        next_button.wait('ready', timeout=20)
        next_button.click()
        logging.info("'Next' button in the 'Select media type' page was clicked")
        next_button.wait('ready', timeout=20)
        next_button.click()
        logging.info("'Next' button in the 'Kernel parameters' page was clicked")
        buildwizard.FXAConfigurationTree1.click()
        buildwizard.send_keystrokes('{SPACE}')
        next_button.click()
        logging.info("'Next' button in the 'Components' page was clicked")
        next_button.wait('ready', timeout=20)
        next_button.click()
        logging.info("'Next' button in the 'Network and remote connection' page was clicked")
        buildwizard.send_keystrokes('{ENTER}')
        logging.info("The first 'YES' button was clicked")
        buildwizard.send_keystrokes('{ENTER}')
        logging.info("The second 'YES' button was clicked")
        buildwizard.send_keystrokes('{DOWN}')
        buildwizard.send_keystrokes('{DOWN}')
        buildwizard.send_keystrokes('{SPACE}')
        fxtext = buildwizard['FXText']
        fxtext.click()
        logging.info("Selecting Text field for entering a key")
        fxtext.send_chars(key_list_f[k])
        fxtext.click()
        key_copy = fxtext.WindowText()
        if key_copy.upper() != key_list_f[k].upper():  # Check if entered key equals to original key
            logging.info("Wrong key is specified: " + str(key_copy) + "\nRetyping key...")
            print("Wrong key is specified: " + str(key_copy))
            fxtext.click()
            fxtext.send_keystrokes('^a')
            fxtext.send_keystrokes('{DELETE}')
            fxtext.send_chars(key_list_f[k])
        else:
            print("Specified license key is correct.")
            logging.info("Specified license key is correct.")
        logging.info(str(key_list_f[k]) + " is selected for media")
        time.sleep(1)
        window.SetFocus()
        fxtext.wait('ready', timeout=20)
        next_button.click()
        logging.info("'Next' button in the 'Enter a license key' page was clicked")
        next_button.wait('ready', timeout=20)
        next_button.click()
        logging.info("'Next' button in the 'Specified license' page was clicked")
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
        next_button.SetFocus()
        next_button.click()
        logging.info("'Next' button in the 'Bootable Media format' page was clicked")
        time.sleep(1)
        buildwizard.FXAFileNameField.click()
        buildwizard.FXAFileNameField.send_keystrokes('^a')
        buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
        logging.info("ISO name field cleaned up")
        time.sleep(1)
        buildwizard.FXAFileNameField.send_chars(new_iso_name)
        iso_name__copy = buildwizard.FXAFileNameField.WindowText()
        if iso_name__copy.lower() != new_iso_name.lower():  # Checking if ISO name was correctly entered
            buildwizard.FXAFileNameField.click()
            buildwizard.FXAFileNameField.send_keystrokes('^a')
            buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
            time.sleep(1)
            logging.info("Re-typing ISO name...")
            buildwizard.FXAFileNameField.send_chars(new_iso_name)
            print("ISO name is correct!")
            logging.info("'" + new_iso_name + "' is a correct ISO name")
        else:
            print("ISO name is correct!")
            logging.info("'" + new_iso_name + "' is a correct ISO name")
        time.sleep(1)
        next_button.SetFocus()
        next_button.wait('exists visible enabled ready active', timeout=20)
        next_button.set_focus()
        next_button.click()
        logging.info("'Next' button in the 'Specify name for ISO' page was clicked")
        next_button.click()
        logging.info("'Next' button in the 'Add drivers' page was clicked")
        next_button.wait('ready')
        next_button.set_focus()
        next_button.click()
        logging.info("'Proceed' button in the 'SUMMARY' was clicked")
        finish_message_box = app.FXAMessageBoxImpl
        finish_message_box.wait("ready active", timeout=50)
        ok_button_in_box = finish_message_box.FXButton
        ok_button_in_box.click()
        logging.info("'OK' button in the message box was clicked")
        if os.path.exists(new_iso_name + ".iso"):  # ISO creation check
            print('Media is created')
            logging.info('Media is created \n====================')
        else:
            print('Media is NOT created')
            logging.info("ISO is not created!")


def start():



    for i in range(len(localization_list)):
        installation(localization_list[i])
        time.sleep(1)
        if if_na:
            print("NA build used")
            main_script(na_keys_list, na_names_list, localization_list[i])
        else:
            print("MAINT build used")
            main_script(keys_list, names_list, localization_list[i])
        time.sleep(1)
        uninstallation()
    print("Operation is complete! See the log file 'C:\MediaCreationLog.log' for more information.")


if __name__ == '__main__':
    start()
