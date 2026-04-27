Why is this so easy? Because the FileSystemObject has a FolderExists method designed to tell you whether or not a folder exists. All you have to do is run the method, passing it the path C:\Scripts. If the method returns True, then the folder was found; if it returns False, then the folder was not found.

Here’s what the code looks like:

Copy
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists("C:\Scripts") Then
    Wscript.Echo "The folder exists."
Else
    Wscript.Echo "The folder does not exist."
End If
OK, but what if you’re looking for a folder on a remote computer; after all, the FileSystemObject is designed to run locally. Well, in that case, just use the WMI class Win32_Directory and look for a folder with the Name C:\\Scripts.

Note. That’s not a misprint. When searching for files and folders using WMI, you must double any \’s found in your query. If you were looking for C:\Scripts\MyScripts\AdminScripts your query would look like this:

Copy
"Select * From Win32_Directory Where Name = " & _
    "'C:\\Scripts\\MyScripts\\AdminScripts'"
This sample script checks for the existence of the folder C:\Scripts on the remote computer atl-ws-01. How do we know if the folder was found? Well, the script reports back the Count (the number of items found). If Count = 0, then C:\Scripts doesn’t exist; if Count = 1, then C:\Scripts does exist.

Copy
strComputer = "atl-ws-01"
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("Select * From Win32_Directory Where " & _
        "Name = 'C:\\Scripts'")
Wscript.Echo colFolders.Count
Ok, now for the hard one: what if you’re looking for a folder named Scripts, but you have no idea whether it’s on drive C or drive D or somewhere else; in other words, you can’t search a specific path like C:\Scripts. That’s all right; you can use WMI to search for a folder with the FileName Scripts (the FileName property is equivalent to what we would call the folder name). Because the entire file system has to be searched this script might take a minute or so to complete (depending on how many folders you have on your machine), but it will do the trick:

Copy
strComputer = "atl-ws-01"
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colFolders = objWMIService.ExecQuery _
    ("Select * From Win32_Directory Where " & _
        "FileName = 'Scripts'")
Wscript.Echo colFolders.Count
This script, by the way, can be used both locally and remotely. Note, too that it’s possible for this script to return more than one item; that’s because all these folders have the FileName Scripts, even though their paths differ:

Copy
C:\Documents and Settings\All Users\Documents\Corporate\Scripts
C:\Scripts
D:\Administrative Tools\Scripts
E:\WMI\Scripts