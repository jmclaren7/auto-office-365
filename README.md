# AutoOffice365
This simple tool works as an interface for the Office Deployment Tool, allowing you to quickly install Microsoft Office without editing the required XML file or using the command line.

<p align="center">
  <img src="https://github.com/jmclaren7/auto-office-365/blob/main/Extras/screenshot1.jpg?raw=true">
</p>

## Features
* Common options like 32bit, adding Access to the install and shared licensing mode are simple check boxes
* Select or type the product ID, channel and build to install
* Select the channel you want and the click "Fetch Versions" to automatically populate a list of build numbers (Thanks to the API at https://office365versions.com)
* Add "[silent]" to the file name to have the installer run automatically with default values
