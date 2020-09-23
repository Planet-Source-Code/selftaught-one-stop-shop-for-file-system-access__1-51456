VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Test()
    Const filename = "C:\znoexist.txt"
    
    Debug.Print PathGetSpecial(sfWINDOWS)
    
    Dim loFileIO As cFileIO: Set loFileIO = New cFileIO
    Dim lsTemp As String
    Dim ldblTemp As Double
    Dim liTemp As Long
    
    With loFileIO
        .FileAccess = GENERIC_WRITE
        .FileCreation = CREATE_ALWAYS
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        .FileShare = FILE_SHARE_READ
        
        Debug.Assert .OpenFile(filename)
        .AppendString "This is a test.  This is only a test."
        .AppendDouble 123456
        .AppendInteger 32000
        .CloseFile
        
        .FileAccess = GENERIC_READ
        .FileCreation = OPEN_EXISTING
        Debug.Assert .OpenFile(filename)
        .GetString lsTemp: Debug.Print lsTemp
        .GetDouble ldblTemp: Debug.Print ldblTemp
        .GetInteger liTemp: Debug.Print liTemp
        .CloseFile
        Debug.Assert FileDelete(filename, False)
    End With
    
    Dim loColl As Collection
    Set loColl = New Collection
    Dim ltFind As tFindFiles
    With ltFind
        .Accessed.High = Date
        .Accessed.Low = Date - 7
        .Recurse = True
        .Path = "C:\"
        .Filter = "*.txt"
        .Size.High = 32767
    End With
    
    Set loColl = FindFiles(ltFind)
    
    Dim lvTemp
    
    Debug.Print "Following are all of the *.txt files on your C:\ drive that were accessed within the last week and are less than 32767 bytes long."
    For Each lvTemp In loColl
        Debug.Print lvTemp
    Next
End Sub
