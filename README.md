AutoUpdate Manifest Creator
===========================

Manifest Creator is used to generate the XML manifest required by the AutoUpdate application when attempting to update a particular program. The manifest is basically a UTF-8 encoded XML recursive file listing of all the files present in the program's startup folder.

Due to IIS not automatically serving all types of files, a further step has been introduced whereby all files are first zipped using 7-zip before being added to the manifest. AutoUpdate Manifest Creator's sister program AutoUpdate automatically unzips the downloaded files when restoring the application directory.

Example Output:
<?xml version="1.0" encoding="UTF-8"?>
<application name="Anime_Collection_Info">
	<file>
		<filename>characters.asp</filename>
		<filepath>\characters.asp</filepath>
		<filesize>4882</filesize>
		<filelastmodified>200501071204</filelastmodified>
	</file>
</application>

------

AutoUpdate Manifest Creator can also be used from the Command Line, using the following 3 input parameters:

 - Autoupdate Manifest Creator.exe "C:\Input Folder" "C:\Output Folder\Manifest.xml" "Application Name"

Parameter 1 is the folder for which the manifest is to be generated, Parameter 2 is the filename to which the manifest is to be saved to and Parameter 3 is the application name to be used in the manifest. Once the operation has completed successfully the application then closes automatically. This is to facilitate batch manifest generation operations.

Created by Craig Lotter, August 2006

*********************************

Project Details:

Coded in Visual Basic .NET using Visual Studio .NET 2005
Implements concepts such as threading and file manipulation.
Level of Complexity: simple

*********************************

Update 20070129.03:

- Creates zipped file manifest (7za encoding)

*********************************

Update 20070402.04:

- Ignores file names that contain either '#' or '&'. 
- Automatically opens folder contained generated manifest files

*********************************

Update 20070403.05:

- Now renames file names that contain either '#' or '&'. Characters are replaced with their HTML code counterpart) 

*********************************

Update 20070514.06:

- Changed the way '#' and '&' characters are handled in file names. Now simply removeds those characters from the zipped file name. File names are restored on extraction
- Fixed bug whereby files of the same name were overwriting one another in the manifest flat folder structure

*********************************

Update 20070614.07:

- Added Command Line capability. Program can be launched with two parameters from the Command Line. First parameter is the folder for which the manifest is to be created. The second parameter is the location where the manifest xml file is to be created. Once the operation is completed, the application closes itself (this is so that it can be used in batch operations).

*********************************

Update 20070619.08:

- Added ManifestUpdated tag to generated XML.
- Added Filepathclear tag to generated XML. This is to deal with the fact that files are now renamed by prefixing a datestamp to their name in order to solve multi-level folder structure bug fixed in update 20070514.06. The clear tag simply records the real path and is used by AutoUpdate to check for existing files.
- Added a third parameter for the Command Line execution: executable.exe Source_Folder Save_File Application_Name

*********************************

Update 20071101.09:

- Fixed bug in handling file names containing # or & characters.
- Improved error handling code.
- Moved from using Microsoft's Settings variables to a plaintext file for config purposes.
