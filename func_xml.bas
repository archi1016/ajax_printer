Attribute VB_Name = "func_xml"
Option Explicit

Public Function XmlAddNodeText(strTag As String, ByVal strText As String) As String
    strText = Replace(strText, vbCr, "")
    strText = Replace(strText, vbLf, "")
    strText = Replace(strText, "&", "&amp;")
    strText = Replace(strText, """", "&quot;")
    strText = Replace(strText, "<", "&lt;")
    strText = Replace(strText, ">", "&gt;")
    XmlAddNodeText = "<" + strTag + ">" + strText + "</" + strTag + ">"
End Function

Public Function XmlAddNodeAttr(strTag As String, ByVal strText As String) As String
    strText = Replace(strText, vbCr, "")
    strText = Replace(strText, vbLf, "")
    strText = Replace(strText, "&", "&amp;")
    strText = Replace(strText, """", "&quot;")
    strText = Replace(strText, "<", "&lt;")
    strText = Replace(strText, ">", "&gt;")
    XmlAddNodeAttr = " " + strTag + "=""" + strText + """"
End Function

Public Function SafeGetXmlNodeText(XmlNode As Object, NodeName As String) As String
    Dim tmpObj As Object
    
    SafeGetXmlNodeText = vbNullString
    If Not (XmlNode Is Nothing) Then
        Set tmpObj = XmlNode.selectSingleNode(NodeName)
        If Not (tmpObj Is Nothing) Then SafeGetXmlNodeText = tmpObj.Text
    End If
    
    Set tmpObj = Nothing
End Function

Public Function SafeGetXmlNodeAttr(XmlNode As Object, NodeAttr As String) As String
    Dim tmpObj As Object
    
    SafeGetXmlNodeAttr = vbNullString
    If Not (XmlNode Is Nothing) Then
        Set tmpObj = XmlNode.Attributes.getNamedItem(NodeAttr)
        If Not (tmpObj Is Nothing) Then SafeGetXmlNodeAttr = tmpObj.Text
    End If
    
    Set tmpObj = Nothing
End Function

Public Function GetXmlObject() As Object
    Set GetXmlObject = CreateObject("MSXML.DOMDocument")
    GetXmlObject.async = False
End Function
