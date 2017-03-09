Attribute VB_Name = "api_reg"
Option Explicit

Public Declare Function RegOpenKeyExW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As Long, _
     ByVal ulOptions As Long, _
     ByVal samDesired As Long, _
     phkResult As Long) As Long

Public Declare Function RegCreateKeyExW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal lpSubKey As Long, _
     ByVal Reserved As Long, _
     ByVal lpClass As Long, _
     ByVal dwOptions As Long, _
     ByVal samDesired As Long, _
     ByVal lpSecurityAttributes As Long, _
     phkResult As Long, _
     ByVal lpdwDisposition As Long) As Long
     
Public Declare Function RegCloseKey Lib "advapi32" _
    (ByVal hKey As Long) As Long

Public Declare Function SHDeleteKeyW Lib "shlwapi" _
    (ByVal hKey As Long, _
     ByVal pszSubKey As Long) As Long
    
Public Declare Function RegDeleteValueW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal lpValueName As Long) As Long

Public Declare Function RegEnumKeyW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal dwIndex As Long, _
     ByVal lpName As Long, _
     ByVal cchName As Long) As Long
     
Public Declare Function RegSetValueExW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal lpValueName As Long, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, ByVal lpData As Long, _
     ByVal cbData As Long) As Long

Public Declare Function RegQueryValueExW Lib "advapi32" _
    (ByVal hKey As Long, _
     ByVal lpValueName As Long, _
     ByVal lpReserved As Long, _
     lpType As Long, _
     ByVal lpData As Long, _
     lpcbData As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_MULTI_SZ = 7

Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_WOW64_64KEY = &H100

Public Declare Function GetPrivateProfileStringW Lib "Kernel32" _
    (ByVal lpAppName As Long, _
     ByVal lpKeyName As Long, _
     ByVal lpDefault As Long, _
     ByVal lpReturnedString As Long, _
     ByVal nSize As Long, _
     ByVal lpFileName As Long) As Long
     
Public Declare Function WritePrivateProfileStringW Lib "Kernel32" _
    (ByVal lpAppName As Long, _
     ByVal lpKeyName As Long, _
     ByVal lpString As Long, _
     ByVal lpFileName As Long) As Long

Public Declare Function GetPrivateProfileSectionNamesW Lib "Kernel32" _
    (ByVal lpszReturnBuffer As Long, _
     ByVal nSize As Long, _
     ByVal lpFileName As Long) As Long
     

