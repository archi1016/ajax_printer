Attribute VB_Name = "define"
Option Explicit

Public Const APPLICATION_ID = "ajax_printer"

Public Const WM_TCP_SERVICE = WM_USER + 137

Public WM_TASKBARCREATED As Long
Public Const TASKBAR_CREATED = "TaskbarCreated"
Public Const WM_TRAYICONCLICK = WM_USER + 168

Public Const AJAX_SERVICE_LISTEN_PORT = "35427"

Public Const PREVIEW_BMP = APPLICATION_ID + ".preview.bmp"
Public Const PREVIEW_PNG = APPLICATION_ID + ".preview.png"

Public Const CMD_SPLIT_CHAR = ","
Public Const CMD_NAME = 0

Public Const CMD_INSTALL = "/INSTALL"
Public Const CMD_UNINSTALL = "/UNINSTALL"

Public Const XML_FILE_HEADER = "<?xml version=""1.0""?>"
Public Const CONTENT_ERROR_404 = "<!DOCTYPE html><html><head><title>404 Not Found</title></head><body><h1>404 Not Found</h1></body></html>"
Public Const ACCESS_CONTROL_ALLOW_HEADERS = "Origin, Content-Type, X-CSRF-TOKEN"

