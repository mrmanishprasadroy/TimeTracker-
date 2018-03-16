Set objFSO = CreateObject("Scripting.FileSystemObject")
' First parameter: original location\file
' Second parameter: new location\file
objFSO.CopyFile "test.mdb", "test_backup.mdb"
'objFSO.DeleteFile("test.mdb")