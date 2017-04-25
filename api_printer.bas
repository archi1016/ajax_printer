Attribute VB_Name = "api_printer"
Option Explicit

Public Declare Function OpenPrinterW Lib "winspool.drv" _
    (ByVal pPrinterName As Long, _
     phPrinter As Long, _
     ByVal pDefault As Long) As Long

Public Declare Function GetPrinterW Lib "winspool.drv" _
    (ByVal hPrinter As Long, _
     ByVal level As Long, _
     ByVal pPrinter As Long, _
     ByVal cbBuf As Long, _
     pcbNeeded As Long) As Long
     
Public Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

'BOOL GetPrinter(
'  _In_  HANDLE  hPrinter,
'  _In_  DWORD   Level,
'  _Out_ LPBYTE  pPrinter,
'  _In_  DWORD   cbBuf,
'  _Out_ LPDWORD pcbNeeded
');

Public Type PRINTER_INFO_2
    pServerName As Long
    pPrinterName As Long
    pShareName As Long
    pPortName As Long
    pDriverName As Long
    pComment As Long
    pLocation As Long
    pDevMode As Long
    pSepFile As Long
    pPrintProcessor As Long
    pDatatype As Long
    pParameters As Long
    pSecurityDescriptor As Long
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

'PRINTER_STATUS_OFFLINE = &H80
