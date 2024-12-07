Function RecursiveGetFiles(strFolder)
  Dim fso, folder, file, files, subfolder
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set folder = fso.GetFolder(strFolder)
  Set files = folder.Files
  For Each file In files
    WScript.Echo file.Path
  Next
  Set subfolders = folder.SubFolders
  For Each subfolder In subfolders
    RecursiveGetFiles subfolder.Path
  Next
End Function

' Example usage:
RecursiveGetFiles "C:\path\to\your\directory"
Set fso = Nothing