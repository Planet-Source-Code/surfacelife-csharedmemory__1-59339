VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSharedMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long

Private Const SECTION_QUERY = &H1
Private Const SECTION_MAP_EXECUTE = &H8
Private Const SECTION_EXTEND_SIZE = &H10
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4

Private Declare Sub copyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Private Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS
Private Const PAGE_READWRITE = &H4

Private address As Long
Private handle As Long


Public Property Let memory(Optional offset As Long, Optional length As Long, value As String)
    Dim b() As Byte
    b = StrConv(value, vbFromUnicode)
    copyMemory ByVal (address + offset), b(0), UBound(b) + 1
End Property

Public Property Get memory(Optional offset As Long, Optional length As Long) As String
    If length > 0 Then
        ReDim b(length - 1) As Byte
        copyMemory b(0), ByVal (address + offset), length
        memory = StrConv(b, vbUnicode)
    End If
End Property

Public Property Let memoryb(Optional offset As Long, Optional length As Long, value() As Byte)
    copyMemory ByVal (address + offset), value(0), UBound(value) + 1
End Property

Public Property Get memoryb(Optional offset As Long, Optional length As Long) As Byte()
    If length > 0 Then
        ReDim b(length - 1) As Byte
        copyMemory b(0), ByVal (address + offset), length
        memoryb = b
    End If
End Property

Public Property Let memoryByte(Optional offset As Long, value As Byte)
    copyMemory ByVal (address + offset), value, 1
End Property

Public Property Get memoryByte(Optional offset As Long) As Byte
    copyMemory memoryByte, ByVal (address + offset), 1
End Property

Public Function sharedMemory(ByVal id As String, Size As Long) As Long
    Dim rv As Long
    If address = 0 Then
        If id <> "" Then
            handle = OpenFileMapping(FILE_MAP_ALL_ACCESS, True, id)
            If handle = 0 And Size > 0 Then handle = CreateFileMapping(&HFFFFFFFF, ByVal 0&, PAGE_READWRITE, 0&, Size + 4, id)
            If handle <> 0 Then rv = MapViewOfFile(handle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
        End If
    End If
    address = rv
End Function

Private Sub Class_Terminate()
    Dim rv As Boolean, handle As Long
    If address <> 0 Then
        rv = UnmapViewOfFile(address)
        rv = CloseHandle(handle)
    End If
End Sub
