"""
Script is created to automate Acronis Linux-based Bootable media creation (ISO) using Acronis Media Builder (MB).
The script can automatically install -> create ISOs -> uninstall MB of existing localization.
How To Use it:
    1. Prepare a VM with Win7 x64 with 1 CD-ROM and no Floppy or flash! It is important step! Disk C: = 100GB
    2. Specify new build number in the variable "build_number". It should look like 50073_
    3. Make sure that na_keys.json and keys.json are present in the root folder near the script. The files with licenses
    4. On the machine, run Setup.bat to set up Python 3 and pywinauto library
    5. Put all big installers of ABR to ./installers folder and run Unzip.bat. MSI will be extracted to separate folders
    6. Run the script:> python ABA11.7_MediaCreation_script.py
    7. Wait when all ISOs of all localizations are created (find them on 'C:\'). DO NOT MOVE THE MOUSE.
DO NOT open any other windows/applications during creating media!
"""

from pywinauto.application import Application
import os
import time
import logging
import json


build_number = "50073_"  # Specify a build number with "_" character in the end. Example: "50064_"
# if_na = True           # Set True if NA build is used

logging.basicConfig(level=logging.INFO, filename='C:\MediaCreationLog.log', format='%(asctime)s %(message)s')


installer_folder_list = []  # Getting a folders list


for (dirpath, dirnames, filenames) in os.walk(".\\installers"):  # Getting a folder list with installers
    installer_folder_list.extend(dirnames)
    break
if installer_folder_list == []:
    print("No MSI installers found. Extracting MSI...")
    logging.info("No MSI installers found. Extracting MSI...")
    os.system(".\\Unzip.bat")
    print("MSI are extracted")
    logging.info("MSI are extracted")
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
    data = json.load(open(".\\keys.json", 'r'))
    keys_list = data["main_keys"]
    trial_us_keys = data["trial_keys"]
except FileNotFoundError:
    print("ERROR: File with keys not found!")
    logging.info("ERROR: File with keys not found!")

try:
    data = json.load(open(".\\na_keys.json", 'r'))
    na_keys_list = data["main_na_keys"]
except FileNotFoundError:
    print("ERROR: File with keys not found!")
    logging.info("ERROR: File with NA keys not found!")

names_list = [
    "AcronisBackupAdvancedWS_11.7_",
    "AcronisBackupAdvancedUniversal_11.7_",
    "AcronisBackupAdvancedHyperV_11.7_",
    "AcronisBackupAdvancedVMware_11.7_",
    "AcronisBackupAdvancedRHEV_11.7_",
    "AcronisBackupAdvancedXEN_11.7_",
    "AcronisBackupAdvancedOracle_11.7_",
    "AcronisBackupEssentials_11.7_",
    "AcronisBackupAdvancedPC_11.7_",
    "AcronisBackupWS_11.7_",
    "AcronisBackupPC_11.7_"
]
trial_us_names = [
    "AcronisBackupAdvancedUniversal_11.7_trial_",
    "AcronisBackupWS_11.7_trial_"
]
na_names_list = [
    'AcronisBackupAdvancedWS_11.7N_',
    'AcronisBackupAdvancedUniversal_11.7N_',
    'AcronisBackupAdvancedHyperV_11.7N_',
    'AcronisBackupAdvancedVMware_11.7N_',
    'AcronisBackupAdvancedRHEV_11.7N_',
    'AcronisBackupAdvancedXEN_11.7N_',
    'AcronisBackupAdvancedOracle_11.7N_',
    'AcronisBackupEssentials_11.7N_',
    'AcronisBackupAdvancedPC_11.7N_',
    'AcronisBackupWS_11.7N_',
    'AcronisBackupPC_11.7N_'
]


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
        print("Something went wrong! See MSI log for more information")


def installation(current_local):
    """ Is called when installation of Media Builder is required """
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        logging.info("An unknown Media Builder is present in system and will be uninstalled.")
        uninstallation()
    print("New Media Builder is being installed. Please wait...")
    logging.info("New Media Builder is being installed.")
    installation_command = u'start /wait msiexec /i C:\\Environment\\installers\\AcronisBackupAdvanced_11.7_' + build_number + current_local + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'  # original ->> u'msiexec /i %CD%\\installers\\AcronisBackupAdvanced_11.7_' + build_number + localization + '\\AcronisBootableComponentsMediaBuilder.msi /quiet /qn'
    os.system(installation_command)
    if os.path.exists('C:\\Program Files (x86)\\Common Files\\Acronis\\MediaBuilder\\MediaBuilder.exe'):
        print("Installation complete!")
        logging.info("Installation complete!")
    else:
        print("Installation FAILED!")
        logging.info("Installation failed!")


def us_keys_extend(trial_us_keys_f):
    """ Extends a non-US list of keys with a US list of keys """
    trial_us_keys_f.extend(trial_us_keys)
    return trial_us_keys_f


def us_names_extend(trial_us_names_f):
    """ Extends a non-US list of names with a US list of names """
    trial_us_names_f.extend(trial_us_names)
    return trial_us_names_f


def main_script(key_list_f, names_list_f, locale):
    """ This function creates ISO """
    for k in range(len(key_list_f)):  # k is an index of a license. This is a loop of creating ISOs
        logging.info("ISO creation starts...")
        new_iso_name = "C:\\" + names_list_f[k] + build_number + locale  # Select a name from list
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
        logging.info("Text field was successfully selected for entering a key")
        fxtext.send_keystrokes(key_list_f[k])  # Select a license key from list
        fxtext.click()
        key_copy = fxtext.WindowText()
        if key_copy.upper() != key_list_f[k].upper():  # Check if entered key equals to original key
            logging.info("Wrong key is specified: " + str(key_copy) + "\nRetyping key...")
            print("Wrong key is specified: " + str(key_copy))
            fxtext.send_keystrokes('^a')
            fxtext.send_keystrokes(key_list_f[k])
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
            print("Localization check: One of these is selected: US, EU, DE, TW, JP, KR, PL, CN, ES")
            logging.info(locale + " localization specified")
        elif locale == "cs-CZ" or locale == "it-IT":
            logging.info(locale + " localization specified")
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
            print("Localization check: One of these is selected: CZ, IT")
        elif locale == "fr-FR":
            logging.info(locale + " localization specified")
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
            time.sleep(1)
            buildwizard.send_keystrokes('{UP}')
            print("Localization check: One of these is selected: FR")
        time.sleep(1)
        next_button.SetFocus()
        next_button.click()
        logging.info("'Next' button in the 'Bootable Media format' page was clicked")
        time.sleep(2)
        buildwizard.FXAFileNameField.click()
        buildwizard.FXAFileNameField.send_keystrokes('^a')
        buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
        time.sleep(2)
        buildwizard.FXAFileNameField.send_chars(new_iso_name)
        iso_name__copy = buildwizard.FXAFileNameField.WindowText()
        if iso_name__copy.lower() != new_iso_name.lower():  # Checking if ISO name was correctly entered
            buildwizard.FXAFileNameField.click()
            buildwizard.FXAFileNameField.send_keystrokes('^a')
            buildwizard.FXAFileNameField.send_keystrokes('{DELETE}')
            time.sleep(2)
            logging.info("Re-typing ISO name...")
            buildwizard.FXAFileNameField.send_chars(new_iso_name)
        time.sleep(2)
        next_button.SetFocus()
        next_button.wait('exists visible enabled ready active', timeout=20)
        logging.info("'" + new_iso_name + "' name specified for new ISO")
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
        finish_message_box.wait("ready active", timeout=30)
        ok_button_in_box = finish_message_box.FXButton
        ok_button_in_box.click()
        logging.info("'OK' button in the message box was clicked")
        if os.path.exists(new_iso_name + ".iso"):  # Check if the created ISO is in the specified location
            print('Media "' + new_iso_name + '.iso" is created')
            logging.info('Media "' + new_iso_name + '.iso" is created ' + "\n====================")
        else:
            print('Media "' + new_iso_name + '.iso" is NOT created')
            logging.info("ISO is not created!" + new_iso_name)



def main():
    for i in range(len(localization_list)):
        print('>> Installation is going to start << ' + localization_list[i])
        installation(localization_list[i])
        time.sleep(5)
        if localization_list[i] == "en-US":
            extended_license_list = us_keys_extend(keys_list)  # extend common keys list with US keys
            extended_names_list = us_names_extend(names_list)  # extend common names list with US names
            main_script(extended_license_list, extended_names_list, localization_list[i])
            time.sleep(5)
        else:
            main_script(keys_list, names_list, localization_list[i])
            time.sleep(5)
        print('>> Uninstallation is going to start <<')
        time.sleep(5)
        uninstallation()
    print("Operation is complete! See the log file 'C:\MediaCreationLog.log' for more information.")


if __name__ == '__main__':
    main()
