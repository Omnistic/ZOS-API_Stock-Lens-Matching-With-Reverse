# ZOS-API Implementation of the Stock Lens Matching tool
* Current version 0.8a
* Developed for OpticStudio (ZEMAX) 21.2.1
* Programing language C#
* User-extension
## Table of content
1. Installation
2. How to use
3. Limitations
## 1. Installation
### Preparing for installation
In OpticStudio, go to **Setup..Project Preferences..Message Boxes** and change the default answer to the second message box (Sample Message: Glass GLASSNAME could not be found ...), and click Ok.

The reason for that is because I think there's a bug in the ZOS-API. When trying to insert lenses made of materials that aren't in the Setup..System Explorer..Material Catalogs, the user is prompted with a warning message:

> Glass GLASSNAME could not be found in the current catalogs. However, ...  
> Click Yes to add this catalog. Click No ...  

The user can then choose between Yes, No, and Cancel. This warning is not raised through the ZOS-API, and makes the extension freeze.

### 1a. Using an executable file
1. Download the files **Reverse_SLM.exe**, and **Reverse_SLM_settings.txt**
2. Copy the files to your **..\Documents\Zemax\ZOS-API\Extensions folder**
This might not work for every computer, and future releases of OpticStudio. If this does not work, try installing by compiling the C# file (see 1b.)

### 1b. By compiling the C# file
1. Download the file **Program.cs**, and **Reverse_SLM_settings.txt**
2. Replace the downloaded **Program.cs** file with the one from a C# User Extension template (**Programing..C#..User Extension** in OpticStudio)
3. Follow the compilation guidelines in [this article](https://my.zemax.com/en-US/Knowledge-Base/kb-article/?ka=KA-01824)
4. Copy the generated executable, and the file **Reverse_SLM_settings.txt** to your **..\Documents\Zemax\ZOS-API\Extensions folder**

## 2. How to use
I don't know yet how to make a settings window in the ZOS-API. Therefore, all the settings are in the file **Reverse_SLM_settings.txt**. There is nearly no error trapping, try to remember this when editing this settings file. Below the list of available settings, their range of values, and what their purpose is.

### Settings
* Reverse = [True or False]: if True, try both orientation of a catalog lens for matching (keep the best orientation only). If False, only the default orientation is used
* Variable = [True or False]: if True, only lenses with at least one variable radius are matched, the others are ignored. If False, all valid lenses are matched
* Vendors = [ALL or VENDOR1, VENDOR2, ...]: only uses catalog lenses from the specified vendors (watch out for spelling mistakes...). ALL means all vendors
* Matches = [UNSIGNED INTEGER]: the number of best matches to display for each lens, and each combination (if it applies)
* EFL/EPD tolerance = [0.0 - 1.0]: is the +/- tolerance on EFL/EPD for corresponding catalog lenses. For example, if the nominal EFL is 100.0 mm, a tolerance of 0.1 will browse all catalog lenses with an EFL between 90.0, and 110.0 mm (I can't be sure if those exact values are included, those are the values I specify as MinEFL, MaxEFL, MinEPD, and MaxEPD in the lens catalog tool of the ZOS-API). This is not given in % as in OpticStudio
* Air thickness compensation = [True or False]: if True, optimize the system using the current Merit Function, and only the present variable air thicknesses. Once again, there isn't much error trapping. If you don't have a Merit Function defined, the value of the Merit Function is always 0. The optimizer is DLS, and it uses the default number of core for the user's computer. If False, replace the catalog lenses without optimization.
* Optimization cycles = [0 or 1 or 5 or 10 or 50]: 0 means the optimizer runs on the automatic number of cycles. 1, 5, 10 or 50 means the optimizer runs for the corresponding number of cycles. It differs from OpticStudio in the sense that the user cannot choose any integer cycles between 0 and 50. This is a limitation of the local optimizer tool in the ZOS-API
* Save best = [True or False]: if True, saves the best combination to FILENAME_SLM_ZOSAPI.ZMX in the same folder as the lens file

Once the settings have been modified, go to OpticStudio and run **Programing..User Extensions..ReverseSLM**.

The results are saved in a file FILENAME_SLM_ZOSAPI_LOG.TXT in the same folder as the lens file.

## 3. Limitations
