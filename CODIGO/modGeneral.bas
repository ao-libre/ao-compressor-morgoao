Attribute VB_Name = "modGeneral"
Option Explicit

Public Const GRAPHIC_PATH = "\GRAFICOS\"
Public Const RESOURCE_PATH = "\RECURSOS\"
Public Const PATCH_PATH = "\PARCHES\"
Public Const EXTRACT_PATH = "\EXTRACCIONES\"

'Public Declare Function GetTickCount Lib "kernel32" () As Long

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function
