Attribute VB_Name = "api"
Option Explicit

Public Const MAX_PATH = 260
Public Const ERROR_SUCCESS = 0
Public Const NO_ERROR = 0
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Const GENERIC_ALL As Long = &H10000000
Public Const GENERIC_EXECUTE  As Long = &H20000000
Public Const GENERIC_READ As Long = &H80000000
Public Const GENERIC_WRITE As Long = &H40000000

Public Const DELETE As Long = &H10000
Public Const READ_CONTROL As Long = &H20000
Public Const WRITE_DAC As Long = &H40000
Public Const WRITE_OWNER As Long = &H80000
Public Const SYNCHRONIZE As Long = &H100000
Public Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE  As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE  As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const SPECIFIC_RIGHTS_ALL As Long = 65535
Public Const ALL_ACCESS As Long = &HF0000 Or SYNCHRONIZE Or &H1FF

Public Declare Function CloseHandle Lib "Kernel32" _
    (ByVal hObject As Long) As Long

Public Declare Sub Sleep Lib "Kernel32" _
    (ByVal dwMilliseconds As Long)

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
    (ByVal lpvDest As Long, _
     ByVal lpvSource As Long, _
     ByVal cbCopy As Long)

Public Declare Function RtlCompareMemory Lib "ntdll" _
    (ByVal lpvSource1 As Long, _
     ByVal lpvSource2 As Long, _
     ByVal nLength As Long) As Long

Public Declare Sub RtlZeroMemory Lib "Kernel32" _
    (ByVal Destination As Long, _
     ByVal nLength As Long)

Public Declare Function CompareStringW Lib "Kernel32" _
    (ByVal Locale As Long, _
     ByVal swCmpFlags As Long, _
     ByVal lpString1 As Long, _
     ByVal cchCount1 As Long, _
     ByVal lpString2 As Long, _
     ByVal cchCount2 As Long) As Long

Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const CSTR_LESS_THAN = 1
Public Const CSTR_EQUAL = 2
Public Const CSTR_GREATER_THAN = 3

Public Declare Function GetWindowsDirectoryW Lib "Kernel32" _
    (ByVal lpBuffer As Long, _
     ByVal uSize As Long) As Long

Public Declare Function GetSystemDirectoryW Lib "Kernel32" _
    (ByVal lpBuffer As Long, _
     ByVal uSize As Long) As Long
     
Public Declare Function GetSystemWow64DirectoryW Lib "Kernel32" _
    (ByVal lpBuffer As Long, _
     ByVal uSize As Long) As Long
     
Public Declare Function GetComputerNameW Lib "Kernel32" _
    (ByVal lpBuffer As Long, _
     lpnSize As Long) As Long

Public Declare Function SetComputerNameExW Lib "Kernel32" _
    (ByVal NameType As Long, _
     ByVal lpBuffer As Long) As Long

Public Const ComputerNameNetBIOS = 0
Public Const ComputerNameDnsHostname = 1
Public Const ComputerNameDnsDomain = 2
Public Const ComputerNameDnsFullyQualified = 3
Public Const ComputerNamePhysicalNetBIOS = 4
Public Const ComputerNamePhysicalDnsHostname = 5
Public Const ComputerNamePhysicalDnsDomain = 6
Public Const ComputerNamePhysicalDnsFullyQualified = 7
Public Const ComputerNameMax = 8

Public Declare Function WaitForSingleObject Lib "Kernel32" _
    (ByVal hHandle As Long, _
     ByVal dwMilliseconds As Long) As Long

Public Declare Function WaitForMultipleObjects Lib "Kernel32" _
    (ByVal nCount As Long, _
     lpHandles As Long, _
     ByVal bWaitAll As Long, _
     ByVal dwMilliseconds As Long) As Long
     
Public Const INFINITE = &HFFFFFFFF
Public Const WAIT_ABANDONED = &H80
Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_ABANDONED_0 As Long = &H80
Public Const WAIT_TIMEOUT = &H102

Public Declare Sub GetLocalTime Lib "Kernel32" _
    (lpSystemTime As SYSTEMTIME)
    
Public Declare Function SetSystemTime Lib "Kernel32" _
    (lpSystemTime As SYSTEMTIME) As Long
    
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Declare Function GetCurrentProcessId Lib "Kernel32" () As Long

Public Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Public Declare Function GetProcessId Lib "Kernel32" _
    (ByVal hProcess As Long) As Long

Public Declare Function TerminateProcess Lib "Kernel32" _
    (ByVal hProcess As Long, _
     ByVal uExitCode As Long) As Long

Public Declare Function GetDriveTypeW Lib "Kernel32" _
    (ByVal lpRootPathName As Long) As Long
    
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NO_ROOT_DIR = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Declare Function GetLogicalDriveStringsW Lib "Kernel32" _
    (ByVal nBufferLength As Long, _
     ByVal lpBuffer As Long) As Long

Public Declare Function GetDiskFreeSpaceExW Lib "Kernel32" _
    (ByVal lpDirectoryName As Long, _
     ByVal lpFreeBytesAvailable As Long, _
     ByVal lpTotalNumberOfBytes As Long, _
     ByVal lpTotalNumberOfFreeBytes As Long) As Long


Public Declare Function GetTempPathW Lib "Kernel32" _
    (ByVal nBufferLength As Long, _
     ByVal lpBuffer As Long) As Long

Public Declare Function GetVolumeInformationW Lib "Kernel32" _
    (ByVal lpRootPathName As Long, _
     ByVal lpVolumeNameBuffer As Long, _
     ByVal nVolumeNameSize As Long, _
     ByVal lpVolumeSerialNumber As Long, _
     ByVal lpMaximumComponentLength As Long, _
     ByVal lpFileSystemFlags As Long, _
     ByVal lpFileSystemNameBuffer As Long, _
     ByVal nFileSystemNameSize As Long) As Long
     
Public Declare Function GetProcessAffinityMask Lib "Kernel32" _
    (ByVal hProcess As Long, _
     lpProcessAffinityMask As Long, _
     lpSystemAffinityMask As Long) As Long
     
Public Declare Function SetProcessAffinityMask Lib "Kernel32" _
    (ByVal hProcess As Long, _
     ByVal dwProcessAffinityMask As Long) As Long


Public Declare Function GlobalMemoryStatusEx Lib "Kernel32" _
    (lpBuffer As MEMORYSTATUSEX) As Long

Public Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

Public Declare Sub GetSystemInfo Lib "Kernel32" _
    (lpSystemInfo As SYSTEM_INFO)
        
Public Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Public Declare Sub GetStartupInfoW Lib "Kernel32" _
    (lpStartupInfo As STARTUPINFO)
    
Public Declare Function CreateProcessW Lib "Kernel32" _
    (ByVal lpApplicationName As Long, _
     ByVal lpCommandLine As Long, _
     ByVal lpProcessAttributes As Long, _
     ByVal lpProcessAttributes As Long, _
     ByVal bInheritHandles As Long, _
     ByVal dwCreationFlags As Long, _
     ByVal lpEnvironment As Long, _
     ByVal lpCurrentDirectory As Long, _
     lpStartupInfo As STARTUPINFO, _
     lpProcessInformation As PROCESS_INFORMATION) As Long

Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Long
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Const STARTF_USESHOWWINDOW = &H1
Public Const STARTF_USESTDHANDLES = &H100

Public Declare Function WideCharToMultiByte Lib "Kernel32" _
    (ByVal codepage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     lpMultiByteStr As Byte, _
     ByVal cbMultiByte As Long, _
     ByVal lpDefaultChar As Long, _
     ByVal lpUsedDefaultChar As Long) As Long

Public Declare Function MultiByteToWideChar Lib "Kernel32" _
    (ByVal codepage As Long, _
     ByVal dwFlags As Long, _
     lpMultiByteStr As Byte, _
     ByVal cbMultiByte As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long) As Long

Public Const CP_UTF8 = 65001

Public Const ByteOrderMark_ANSI = 0
Public Const ByteOrderMark_UNICODE = 1
Public Const ByteOrderMark_UTF8 = 2

Public Declare Function GetVersionExW Lib "Kernel32" _
    (lpVersionInfo As OSVERSIONINFOEX) As Long
    
Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(255) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Const WINDOWS_VERSION_XP = 0
Public Const WINDOWS_VERSION_VISTA = 1
Public Const WINDOWS_VERSION_7 = 2
Public Const WINDOWS_VERSION_8 = 3

Public Declare Function LocalAlloc Lib "Kernel32" _
    (ByVal uFlags As Long, _
     ByVal uBytes As Long) As Long

Public Declare Function LocalFree Lib "Kernel32" _
    (ByVal hMem As Long) As Long

Public Declare Function GetVolumeNameForVolumeMountPointW Lib "Kernel32" _
    (ByVal lpszVolumeMountPoint As Long, _
     ByVal lpszVolumeName As Long, _
     ByVal cchBufferLength As Long) As Long

Public Declare Function LoadLibraryW Lib "Kernel32" _
    (ByVal lpFileName As Long) As Long

Public Declare Function LoadLibraryExW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal hFile As Long, _
     ByVal dwFlags As Long) As Long
    
Public Const LOAD_LIBRARY_AS_DATAFILE = &H2

Public Declare Function FreeLibrary Lib "Kernel32" _
    (ByVal hModule As Long) As Long
    
Public Declare Function FindResourceW Lib "Kernel32" _
    (ByVal hModule As Long, _
     ByVal lpName As Long, _
     ByVal lpType As Long) As Long

Public Const RT_CURSOR = "#1"
Public Const RT_BITMAP = "#2"
Public Const RT_ICON = "#3"
Public Const RT_MENU = "#4"
Public Const RT_DIALOG = "#5"
Public Const RT_STRING = "#6"
Public Const RT_FONTDIR = "#7"
Public Const RT_FONT = "#8"
Public Const RT_ACCELERATOR = "#9"
Public Const RT_RCDATA = "#10"
Public Const RT_MESSAGETABLE = "#11"
Public Const RT_GROUP_CURSOR = "#12"
Public Const RT_GROUP_ICON = "#14"
Public Const RT_VERSION = "#16"
Public Const RT_DLGINCLUDE = "#17"
Public Const RT_PLUGPLAY = "#19"
Public Const RT_VXD = "#20"
Public Const RT_ANICURSOR = "#21"
Public Const RT_ANIICON = "#22"

Public Declare Function SizeofResource Lib "Kernel32" _
    (ByVal hModule As Long, _
     ByVal hResInfo As Long) As Long
     
Public Declare Function LoadResource Lib "Kernel32" _
    (ByVal hModule As Long, _
     ByVal hResInfo As Long) As Long
     
Public Declare Function LockResource Lib "Kernel32" _
    (ByVal hglbResource As Long) As Long
    

Public Declare Function FreeResource Lib "Kernel32" _
    (ByVal hglbResource As Long) As Long

Public Declare Function CreateEventW Lib "Kernel32" _
    (ByVal lpEventAttributes As Long, _
     ByVal bManualReset As Long, _
     ByVal bInitialState As Long, _
     ByVal lpName As Long) As Long
     
Public Declare Function ResetEvent Lib "Kernel32" _
    (ByVal hEvent As Long) As Long

Public Declare Function SetEvent Lib "Kernel32" _
    (ByVal hEvent As Long) As Long

Public Declare Function lstrcpyW Lib "Kernel32" _
    (ByVal lpString1 As Long, _
     ByVal lpString2 As Long) As Long
