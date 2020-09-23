Attribute VB_Name = "MyBas"
'Thank You for downloading my FileSearching entry

'Please email me your Questions on anything in here or comments as well

'My E-Mail = Mike_3D@hotmail.com



Option Explicit

Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const MAX_PATH = 260

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80




