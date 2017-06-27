Option Explicit
Dim PathHere 
Dim FSO
Dim i, j

Dim DestFolder 
DestFolder = "D:\Temp\" ' Destination Folder
Dim re
Set re = new regexp

re.Pattern = "-[0-9]{3}"
re.IgnoreCase = True
re.Global = True


SET FSO = CreateObject("Scripting.FileSystemObject") 
PathHere = FSO.GetAbsolutePathName(".")

On Error Resume Next

Dim FileExt
FileExt = Array("mp4", "avi")

Dim sFolder
For Each sFolder In FSO.GetFolder(PathHere).Subfolders
	IF re.Test(sFolder) Then
		For i = 0 To UBound(FileExt)
		
			FSO.MoveFile sFolder & "\*." & FileExt(i), DestFolder
		
			IF FSO.FolderExists(sFolder) AND NOT FSO.FileExists(sFolder & "\*." & FileExt(i)) THEN
				FSO.DeleteFolder sFolder
			End If 
		Next
	END IF
Next


SET re = Nothing
SET FSO = Nothing