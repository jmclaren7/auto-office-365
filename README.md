# AutoOffice365
This simple tool works as an interface for the Office Deployment Tool, allowing you to quickly install Microsoft Office without editing required XML file or using the command line.

<p align="center">
  <img src="https://github.com/jmclaren7/auto-office-365/blob/main/Extras/screenshot1.jpg?raw=true">
</p>

## Features
* Common options like 32bit, adding Access to the install and shared licensing mode are simple check boxes
* Select or type the product ID, channel and build to install
* Select the channel you want and the click "Fetch Versions" to automatically populate a list of build numbers (Thanks to the API at https://office365versions.com)

## Do Do
* Add an option to specify where the install data is saved (DownloadPath)
* Add an option to use already downloaded install data
* Add language selection with support for multiple language selection
* Add option to hide log, ODT console and ODT window
* Add more product ID options to make individual product install possible with support for multiple product selection and product exclusion
* Add activation options
* Add DeviceBasedLicensing and SCLCache options
* Improve download progress indicator

## A Note About AutoIT and Antivirus
AutoIt is an easy to use and capable scripting language that is also easily compiled to an executable. It's also easy for AntiVirus software to falsely identify all AutoIT scripts as malware because all compiled scripts contain the same AutoIT interpreter. I make every effort to reduce false positives but detections are inevitable, please let me know if you experience frequent false positives with any popular antivirus software.
