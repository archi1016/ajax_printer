Attribute VB_Name = "api_path"
Option Explicit

Public Declare Function PathCombineW Lib "Shlwapi" _
    (ByVal pszPathOut As Long, _
     ByVal pszPathIn As Long, _
     ByVal pszMore As Long) As Long

