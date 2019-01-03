VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoryDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const strurl = "http://www.bwlc.net/bulletin/prevtrax.html?page=XX"
Public Sub GetHistoryData(pageno As Long, txtarr() As ExpectData)
    On Error Resume Next
    Dim httpreq As New XMLHTTP60 ' MSXML2.XMLHTTP30
    Dim str As String
    httpreq.Open "get", Replace(strurl, "XX", pageno), False
    'HttpReq.setRequestHeader "contentType", "image/jpeg"
    httpreq.send
    DoEvents
    If httpreq.Status <> 200 Then
        'MsgBox HttpReq.Status
        ReDim txtarr(0)
        Exit Sub
    End If
    strHtml = httpreq.responseText
''''    Dim FilePath As String
''''    FilePath = ThisWorkbook.Path & "\test.txt"
''''    Kill FilePath
''''    Dim fso As New FileSystemObject
''''    Dim fo As TextStream
''''    fso.CreateTextFile FilePath
''''    Set fo = fso.OpenTextFile(FilePath, ForWriting)
''''    fo.Write strHtml
''''    fo.Close
    sp = InStr(strHtml, "<table class")
    ep = InStr(sp, strHtml, "</table>")
    txt = Mid(strHtml, sp, ep - sp - 1)
    sp = InStr(txt, "<tr")
    txt = Mid(txt, sp, Len(txt) - sp)
    txt = Replace(Replace(Replace(txt, Chr(13), ""), Chr(10), ""), Chr(9), "")
    'txt = Replace(txt, " ", "")
    'txt = Replace(txt, " ", "")
    'txt = Replace(txt, "    ", "")
    txt = Replace(txt, "odd", "")
    txt = Replace(txt, "class=""""", "")
    
    trarr = Split(txt, "</tr><tr >")
    'Dim txtarr() As ExpectData
    ReDim txtarr(UBound(trarr))
    For i = 1 To UBound(trarr)
        txttr = Replace(trarr(i), "</tr>", "")
        Set txtarr(i) = New ExpectData
        txttdarr = Split(txttr, "</td><td>")
        txtarr(i).expect = Replace(txttdarr(0), "<td>", "")
        txtarr(i).OpenCode = txttdarr(1)
        txtarr(i).OpenTime = Replace(txttdarr(2), "</td>", "")
    Next
    
End Sub

