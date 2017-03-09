Attribute VB_Name = "api_file"
Option Explicit

Public Const ERROR_IO_PENDING As Long = 997

Public Const INVALID_HANDLE_VALUE = -1
Public Const INVALID_FILE_ATTRIBUTES = -1
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_NO_BUFFERING As Long = &H20000000
Public Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Public Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_DELETE = &H4

Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_ALWAYS = 4
Public Const OPEN_EXISTING = 3
Public Const TRUNCATE_EXISTING = 5

Public Declare Function SetCurrentDirectoryW Lib "Kernel32" _
    (ByVal lpPathName As Long) As Long

Public Declare Function RemoveDirectoryW Lib "Kernel32" _
    (ByVal lpPathName As Long) As Long
     
Public Declare Function CreateDirectoryW Lib "Kernel32" _
    (ByVal lpPathName As Long, _
     ByVal lpSecurityAttributes As Long) As Long
     
Public Declare Function CreateFileW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal dwDesiredAccess As Long, _
     ByVal dwShareMode As Long, _
     ByVal lpSecurityAttributes As Long, _
     ByVal dwCreationDisposition As Long, _
     ByVal dwFlagsAndAttributes As Long, _
     ByVal hTemplateFile As Long) As Long
     
Public Declare Function WriteFile Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuffer As Long, _
     ByVal nNumberOfBytesToWrite As Long, _
     lpNumberOfBytesWritten As Long, _
     ByVal lpOverlapped As Long) As Long
     
Public Declare Function ReadFile Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuffer As Long, _
     ByVal nNumberOfBytesToRead As Long, _
     lpNumberOfBytesRead As Long, _
     ByVal lpOverlapped As Long) As Long
       
Public Declare Function GetFileSize Lib "Kernel32" _
    (ByVal hFile As Long, _
     lpFileSizeHigh As Long) As Long

Public Declare Function GetFileSizeEx Lib "Kernel32" _
    (ByVal hFile As Long, _
     lpFileSize As Currency) As Long
     
Public Declare Function GetFileAttributesW Lib "Kernel32" _
    (ByVal lpFileName As Long) As Long
    
Public Declare Function SetFileAttributesW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     ByVal dwFileAttributes As Long) As Long

Public Declare Function GetFileTime Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal lpCreationTime As Long, _
     ByVal lpLastAccessTime As Long, _
     ByVal lpLastWriteTime As Long) As Long
     
Public Declare Function SetFileTime Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal lpCreationTime As Long, _
     ByVal lpLastAccessTime As Long, _
     ByVal lpLastWriteTime As Long) As Long
    
Public Declare Function SetEndOfFile Lib "Kernel32" _
    (ByVal hFile As Long) As Long
    
Public Declare Function SetFilePointerEx Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal liDistanceToMove As Currency, _
     lpNewFilePointer As Currency, _
     ByVal dwMoveMethod As Long) As Long

Public Declare Function DeleteFileW Lib "Kernel32" _
    (ByVal lpFileName As Long) As Long
    
Public Declare Function CopyFileW Lib "Kernel32" _
    (ByVal lpExistingFileName As Long, _
     ByVal lpNewFileName As Long, _
     ByVal bFailIfExists As Long) As Long
     
Public Declare Function CopyFileExW Lib "Kernel32" _
    (ByVal lpExistingFileName As Long, _
     ByVal lpNewFileName As Long, _
     ByVal lpProgressRoutine As Long, _
     ByVal lpData As Long, _
     ByVal pbCancel As Long, _
     ByVal dwCopyFlags As Long) As Long

Public Const COPY_FILE_ALLOW_DECRYPTED_DESTINATION = &H8
Public Const COPY_FILE_COPY_SYMLINK = &H800
Public Const COPY_FILE_FAIL_IF_EXISTS = &H1
Public Const COPY_FILE_NO_BUFFERING = &H1000
Public Const COPY_FILE_OPEN_SOURCE_FOR_WRITE = &H4
Public Const COPY_FILE_RESTARTABLE = &H2

Public Const PROGRESS_CANCEL = 1
Public Const PROGRESS_CONTINUE = 0
Public Const PROGRESS_QUIET = 3
Public Const PROGRESS_STOP = 2

Public Declare Function MoveFileW Lib "Kernel32" _
    (ByVal lpExistingFileName As Long, _
     ByVal lpNewFileName As Long) As Long

Public Declare Function FindFirstFileW Lib "Kernel32" _
    (ByVal lpFileName As Long, _
     lpFindFileData As WIN32_FIND_DATA) As Long
     
Public Declare Function FindNextFileW Lib "Kernel32" _
    (ByVal hFindFile As Long, _
     lpFindFileData As WIN32_FIND_DATA) As Long
     
Public Declare Function FindClose Lib "Kernel32" _
    (ByVal hFindFile As Long) As Long

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(519) As Byte
    cAlternate(27) As Byte
End Type

Public Declare Function GetLogicalDrives Lib "Kernel32" () As Long


Public Declare Function SetFileValidData Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal ValidDataLength As Currency) As Long


Public Declare Function GetOverlappedResult Lib "Kernel32" _
    (ByVal hFile As Long, _
     ByVal lpOverlapped As Long, _
     lpNumberOfBytesTransferred As Long, _
     ByVal bWait As Long) As Long

Public Declare Function FileTimeToLocalFileTime Lib "Kernel32" _
    (lpFileTime As FILETIME, _
     lpLocalFileTime As FILETIME) As Long

Public Declare Function FileTimeToSystemTime Lib "Kernel32" _
    (lpFileTime As FILETIME, _
     lpSystemTime As SYSTEMTIME) As Long

Public Declare Function CompareFileTime Lib "Kernel32" _
    (lpFileTime1 As FILETIME, _
     lpFileTime2 As FILETIME) As Long


Public Const FILE_LIST_DIRECTORY = 1

Public Declare Function ReadDirectoryChangesW Lib "Kernel32" _
    (ByVal hDirectory As Long, _
     ByVal lpBuffer As Long, _
     ByVal nBufferLength As Long, _
     ByVal bWatchSubtree As Long, _
     ByVal dwNotifyFilter As Long, _
     ByVal lpBytesReturned As Long, _
     ByVal lpOverlapped As Long, _
     ByVal lpCompletionRoutine As Long) As Long

Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Public Const FILE_NOTIFY_CHANGE_NAME = &H3
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20
Public Const FILE_NOTIFY_CHANGE_CREATION = &H40
Public Const FILE_NOTIFY_CHANGE_EA = &H80
Public Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Public Const FILE_NOTIFY_CHANGE_STREAM_NAME = &H200
Public Const FILE_NOTIFY_CHANGE_STREAM_SIZE = &H400
Public Const FILE_NOTIFY_CHANGE_STREAM_WRITE = &H800
Public Const FILE_NOTIFY_VALID_MASK = &HFFF

Public Type FILE_NOTIFY_INFORMATION
    NextEntryOffset As Long
    Action As Long
    FileNameLength As Long
    'FileNameBinary As Long
End Type

Public Const FILE_ACTION_ADDED = &H1
Public Const FILE_ACTION_REMOVED = &H2
Public Const FILE_ACTION_MODIFIED = &H3
Public Const FILE_ACTION_RENAMED_OLD_NAME = &H4
Public Const FILE_ACTION_RENAMED_NEW_NAME = &H5
Public Const FILE_ACTION_ADDED_STREAM = &H6
Public Const FILE_ACTION_REMOVED_STREAM = &H7
Public Const FILE_ACTION_MODIFIED_STREAM = &H8

