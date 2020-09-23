Attribute VB_Name = "Module1"
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE

Public Const FILE_MAP_COPY = SECTION_QUERY
Public Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Public Const FILE_MAP_READ = SECTION_MAP_READ
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_WRITECOPY = &H8

Public Const ERROR_ALREADY_EXISTS = 183&

Public ptrStat As Long
Public ptrShare As Long
Public hFile As Long
Public lngStore() As Long



Public Sub PutMemoryData(strdata As String)
Dim a As Long
Dim x As Long

    a = Len(strdata)
    
    ReDim lngStore(a)

    For x = 0 To a - 1
        lngStore(x) = Asc(Mid(strdata, x + 1, 1))
    Next
    
    MsgBox ptrShare & " " & a
    
    CopyMemory ByVal ptrShare, a, 4
    CopyMemory ByVal (ptrShare + 4), lngStore(0), a * 4
End Sub



Public Function GetMemoryData() As String
Dim a As Long
Dim x As Long
Dim strdata As String

CopyMemory a, ByVal ptrShare, 4

ReDim lngStore(a)

    CopyMemory lngStore(0), ByVal (ptrShare + 4), a * 4

        For x = 0 To a
            strdata = strdata & Chr(lngStore(x))
        Next

    GetMemoryData = strdata
End Function



    
Function OpenMemory() As Boolean
Dim strName As String
Dim e As Long


strName = "TaroFTP"

hFile = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0&, 65535, "TaroFTP")

e = Err.LastDllError
MsgBox e & " " & hFile

If hFile Then

    ptrShare = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0&, 0&, 0&)
        If ptrShare <> 0 Then

                If e <> ERROR_ALREADY_EXISTS Then
                    MsgBox "openedit"
                    CopyMemory ByVal ptrShare, 0, 4
                End If
        Else
            MsgBox "Unable to map view of memory"
            OpenMemory = False
            Exit Function
        End If
Else
    MsgBox "Unable to get memory map handle."
    OpenMemory = False
    Exit Function
End If
'CloseHandle hFile
OpenMemory = True
End Function

Public Sub CloseMemory()
UnmapViewOfFile ptrShare
End Sub

