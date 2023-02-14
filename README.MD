# INTRODUCTION
This script uninstalls (if present) "Microsoft Access database engine 2010 (English)" and then performs a "quick repair" of O365 (32-bit)

# AIM
This script was created because A360 latest versions (v26-v27) support the usage of the latest access db engine (for using Excel as DB commands):
"Microsoft Access Database Engine 2016 Redistributable", which comes installed by default with 32-bit O365 installation

# ADDITIONAL INFO

* When installing "Microsoft Access database engine 2010 (English)", this overwrites the "Microsoft Access Database Engine 2016 Redistributable" that is installed
by 32-bit O365
* When uninstalling "Microsoft Access database engine 2010 (English)" this removes all trace of the 32-bit odbc driver for using excel as db
* After uninstalling "Microsoft Access database engine 2010 (English)", it is necessary to perform a quick repair in O365 so that the 2016 32-bit odbc driver is present again under
"Drivers" tab in ODBC Data Sources (32-bit)
