ZipFolderEx
======================

A Brief Intro
-------------------

ZipFolderEx is a utility that extends the built-in ZIP compress functionality of Windows.
With ZipFolderEx, you can easily to unzip files without third-party softwares.

Screenshot:

![zipfolderex](https://cloud.githubusercontent.com/assets/6334728/6682350/72c6047e-ccdc-11e4-8311-8cc1b7c6eddd.png)


This project is base on the [CppShellExtContextMenuHandler](https://code.msdn.microsoft.com/windowsdesktop/cppshellextcontextmenuhandl-410a709a) project on Microsoft Dev Center.

References:

1. [Themed menu’s icons, a complete Vista and XP solution](http://www.nanoant.com/programming/themed-menus-icons-a-complete-vista-xp-solution)
2. [How to zip/unzip files](https://vcpptips.wordpress.com/tag/copyhere/)


Usage
-------------------

(Run cmd as Administrator)

Registering and unregistering .DLL: (64bit version is recommended)

regsvr32 ZipFolderEx.dll		(for Windows 32bit)
**regsvr32 ZipFolderEx64.dll		(for Windows 64bit)**

regsvr32 ZipFolderEx.dll /u		(for Windows 32bit)
**regsvr32 ZipFolderEx64.dll /u	(for Windows 64bit)**

Note: After unregister the DLL, you might need to log off then log in to delete the DLL file, this is a windows behavior.

Version History
-------------------
* v0.1 First working version
  1. Add "Extract Here" and "Extract to XXX" menu items in Shell Context Menu for Zip files.
  2. Use Zip file associated icon for these 2 menu items.
* v0.2 minor update
  1. Targeting Visual Studio 2017 xp framework.
  2. Add/Rmove "Extract All..." default menu item as required.
