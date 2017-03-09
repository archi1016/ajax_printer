Attribute VB_Name = "api_ole"
Option Explicit

Public Declare Function CLSIDFromString Lib "ole32" _
    (ByVal str As Long, id As GUID) As Long

