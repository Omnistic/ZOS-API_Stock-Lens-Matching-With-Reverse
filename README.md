# ZOS-API Implementation of the Stock Lens Matching tool
* Current version 0.8a
* Developed for OpticStudio (ZEMAX) 21.2.1
* Programing language C#
* User-extension
## Table of content
1. Installation
2. How to use
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



## How to use
