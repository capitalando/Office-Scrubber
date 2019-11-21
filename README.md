### **Using the Offscrub scripts**

The Offscrub vbs scripts can be used to automate the removal of Office products. The scripts will uninstall the existing Office products regardless of the current health of the installation. The Remove-PreviousOfficeInstalls.ps1 script will determine which version of Office is currently installed and will call the appropriate Offscrub vbs script to remove the Office products installations.

These Offscrub vbs scripts have had the defaults set to run silently, force quit any open copies of office products, and not prompt the user
These settings are needed for an automation project I'm using.


The Offscrub vbs files included are:

* **OffScrub03.vbs** - Used to remove Office 2003 products.
* **OffScrub07.vbs** - Used to remove Office 2007 products.
* **OffScrub10.vbs** - Used to remove Office 2010 products.
* **OffScrub_O15msi.vbs** - Used to remove Office 2013 MSI products.
* **OffScrub_O16msi.vbs** - Used to remove Office 2016 MSI products.
* **OffScrubc2r.vbs** - Used to remove Click-To-Run products.

More information can be found at: https://blogs.technet.microsoft.com/odsupport/2011/04/08/how-to-obtain-and-use-offscrub-to-automate-the-uninstallation-of-office-products/
