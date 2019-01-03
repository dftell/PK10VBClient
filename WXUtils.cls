VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WXUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public url_Params
Public Message
Public skey
Public wxsid
Public wxuin
Public pass_ticket
Public isgrayscale
Public SyncKey
Public syncCheckKey
Public uuid
Public SelfBean As WXUser
Public FriendList As Dictionary
Public CurrData As Dictionary
Public webwx_data_ticket
Const pic_url = "https://wx2.qq.com/cgi-bin/mmwebwx-bin/webwxgetmsgimg"
Const pic_url_qq = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxgetmsgimg"
Const vidoe_url = "https://wx2.qq.com/cgi-bin/mmwebwx-bin/webwxgetvideo"
Const vidoe_url_qq = "https://wx.qq.com/cgi-bin/mmwebwx-bin/webwxgetvideo"
Public cookie As String  '= "pgv_pvi=6709982208 webwxuvid=eac6c967c1de446c73c08acc64904546218df93ddfedacbf78aeb8150ab77b8f76c2abd11d782cf5b59b46ffb9ca213e pgv_si=s1867008000 "
Const cRunWhat = "refreshCode"
Const cRunIntervalSeconds = 300
Public paramData As String
Public iCheckLoginTip As Long
Public strUserUrl As String
Public strStamp As String
Public userList As Dictionary
Public DicFriends As Dictionary
Public strHostFlg As String
Public ScanedTheQCore As Boolean
Dim MyAllContactors() As WXContactor
'Public strLoginFlg As String
'Public strLoginStatus As String
Const cImgPath = "\test.bmp"
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function getMyAllContactors() As WXContactor()
    getMyAllContactors = MyAllContactors
End Function



Function getWx_Code(img As Image) As String
    Dim uuid As String
    uuid = getUUID()
    'Sheet10.Cells(4, 7) = uuid
    If Not getQR_Code(uuid, cImgPath) Then
        Exit Function
    End If
    Dim filePath As String
    filePath = App.Path & cImgPath
    With img
        .Stretch = True
        .Picture = LoadPicture(filePath)
        
        '.PictureSizeMode = fmPictureSizeModeStretch
    End With
    iCheckLoginTip = 1
    'Sheet10.Cells(4, 6) = iCheckLoginTip
End Function

Function getUUID() As String
    Dim httpreq As New XMLHTTP60
    strStamp = getTimeStap()
    'Sheet10.Cells(4, 8) = strStamp
    strurl = "https://login.weixin.qq.com/jslogin?appid=wx782c26e4c19acffb&redirect_uri=https%3A%2F%.2Fwx.qq.com%2Fcgi-bin%2Fmmwebwx-bin%2Fwebwxnewloginpage&fun=new&lang=zh_CN&_=" & strStamp
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "contentType", "text/html;charset=UTF-8"
    httpreq.send
    Dim retstr As String
    retstr = httpreq.responseText
    If httpreq.Status <> 200 Then
        getUUID = ""
        Exit Function
    End If
    Dim retArr() As String
    retArr = Split(retstr, ";")
    stralluuid = retArr(1)
    Dim uuidArr() As String
    uuidArr = Split(stralluuid, """")
    getUUID = uuidArr(1)
End Function
Function getQR_Code(uuid As String, ImgPath As String) As Boolean
    On Error Resume Next
    Dim ret As Image
    If uuid = "" Then
        Exit Function
    End If
    Dim strurl As String
    Dim httpreq As New XMLHTTP60
    strurl = "https://login.weixin.qq.com/qrcode/" & uuid
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "contentType", "image/jpeg"
    httpreq.send
    strheader = httpreq.getResponseHeader("Content-Type")
    Dim ImgLen As Long
    ImgLen = httpreq.getResponseHeader("Content-Length")
    If httpreq.Status <> 200 Then
        'MsgBox httpreq.Status
        Dim test As String
        test = ""
        Exit Function
    End If
    
    
    Dim filePath As String
    filePath = App.Path & ImgPath
    Kill filePath
    Dim stm As New Stream
    stm.Type = adTypeBinary
    stm.Mode = 3
    stm.Open
    Err.Clear
    stm.Write httpreq.responseBody
    If Err.Number <> 0 Then
        stm.Close
        'MsgBox Err.Description
        Exit Function
    End If
    If stm.Size = 0 Then Exit Function
    Err.Clear
    stm.SaveToFile filePath
    If Err.Number <> 0 Then
        stm.Close
        MsgBox Err.Description
        Exit Function
    End If
    httpreq.abort
    stm.Close

    'Kill filePath
    getQR_Code = True
End Function

Function wait_login(ByRef strReURL As String) As Boolean
    Dim strurl As String
    'Sheet10.Cells(3, 11) = ""
    'Sheet10.Cells(4, 10) = Now()
    'iCheckLoginTip = Sheet10.Cells(4, 6).Text
    'uuid = Sheet10.Cells(4, 7).Text
    'strTimeStamp = Sheet10.Cells(4, 8).Text
    strurl = "https://login.weixin.qq.com/cgi-bin/mmwebwx-bin/login?tip=" & iCheckLoginTip & "&uuid=" & uuid + "&_=" & strTimeStamp
    'Sheet10.Cells(1, 5) = strUrl
    Dim httpreq As New XMLHTTP60
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "contentType", "text/html;charset=UTF-8"
    httpreq.send
    'strheader = HttpReq.getAllResponseHeaders()
    strheader = httpreq.getResponseHeader("Content-Type")
    Dim ImgLen As Long
    ImgLen = httpreq.getResponseHeader("Content-Length")
    If httpreq.Status <> 200 Then
        'MsgBox HttpReq.Status
        Exit Function
    End If
    'MsgBox HttpReq.responseBody
    Dim strHtml As String
    strHtml = httpreq.responseText
    'Sheet10.Cells(2, 11) = strHtml
    Dim strflg As String
    Dim strstatus As String
    Dim strReDirUrl As String
    strstatus = Split(strHtml, ";")(0)
    strflg = Split(strstatus, "=")(1)
    'Sheet10.Cells(2, 11) = strflg
    'strLoginFlg = strflg
    If strflg <> "200" Then
        If strflg = "408" Then
            'Sheet10.Cells(3, 11) = "未扫描"
            iCheckLoginTip = 1
        End If
        If strflg = "201" Then
            'Sheet10.Cells(3, 11) = "未登陆"
            iCheckLoginTip = 0
        End If
        'Sheet10.Cells(4, 6) = iCheckLoginTip
        'strLoginStatus = iCheckLoginTip
        Exit Function
    End If
    'Sheet10.Cells(5, 1) = ""
    strReDirUrl = Split(strHtml, ";")(1)
    strReURL = strReDirUrl
    If InStr(strReDirUrl, "wx2.") > 1 Then
        'Sheet10.Cells(5, 1) = "2"
        strHostFlg = "2"
    End If
    'strHostFlg = Sheet10.Cells(5, 1)
    strUserUrl = Split(strReDirUrl, """")(1)
    strReURL = strUserUrl
    url_Params = strUserUrl
    'Sheet10.Cells(3, 11) = "已登陆"
    'Sheet10.Cells(2, 5) = strUserUrl
    wait_login = True
End Function

Function getCookie(strurl As String) As Boolean
    On Error Resume Next
    'Dim strUrl As String
    strurl = strUserUrl & "&fun=new"
    Dim httpreq As New XMLHTTP60
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "contentType", "text/html;charset=UTF-8"
    httpreq.send
    Dim strCookie As String
    If httpreq.Status <> 200 Then
        MsgBox httpreq.Status
        Exit Function
    End If
    Dim xmldoc As New DOMDocument60
    xmldoc.LoadXML httpreq.responseText
    Dim xmlnodes As IXMLDOMNodeList
    Dim retnode As IXMLDOMNode
    Set retnode = xmldoc.SelectSingleNode("/error/ret")
    If retnode Is Nothing Or retnode.Text = 1 Then
        MsgBox "重链接回复错误！"
        Exit Function
    End If
    Set xmlnodes = xmldoc.SelectNodes("/error/*")
    skey = xmlnodes(2).Text
    wxsid = xmlnodes(3).Text
    wxuin = xmlnodes(4).Text
'''    For I = 1 To 4
'''        Sheet10.Cells(1, 11 + I) = xmlnodes(I + 1).nodeName
'''        Sheet10.Cells(2, 11 + I) = xmlnodes(I + 1).Text
'''    Next
    pass_ticket = xmlnodes(5).Text
    strCookie = httpreq.getResponseHeader("Set-Cookie")
    Dim cookieArr() As String
    cookieArr = Split(strCookie, ";")
'    Sheet10.Cells(3, 11) = strCookie
    scookie = scookie & "wxpluginkey=" & getTimeStap() + "; "
    scookie = scookie & cookieArr(0) & ";"
    scookie = scookie & "wxsid=" & wxsid & ";"
    scookie = scookie & "mm_lang=zh_CN; MM_WX_NOTIFY_STATE=1; MM_WX_SOUND_STATE=1;"
    cookie = scookie
'    Sheet10.Cells(1, 1) = Cookie
    getCookie = True
End Function

Function initSelf() As Boolean
    On Error Resume Next
    Dim postUrl As String
''    strStamp = Sheet10.Cells(4, 8)
''    strHostFlg = Sheet10.Cells(5, 1)
    postUrl = "https://wx" & strHostFlg & ".qq.com/cgi-bin/mmwebwx-bin/webwxinit?r=" & getTimeStap() & "&lang=zh_CN&pass_ticket=" & pass_ticket
    '{"BaseRequest":{"Uin":"UUUU","Sid":"SSSS","Skey":"KKKK","DeviceID":"e123456789012345"}}
    Dim strModel As String
    Dim strPost As String
    strModel = "{""BaseRequest"":{""Uin"":""UUUU"",""Sid"":""SSSS"",""Skey"":""KKKK"",""DeviceID"":""eDDDD""}}"
    strPost = Replace(strModel, "UUUU", wxuin)
    strPost = Replace(strPost, "SSSS", wxsid)
    strPost = Replace(strPost, "KKKK", skey)
    strPost = Replace(strPost, "DDDD", Right("00000000000" & strStamp, 15))
    paramData = strPost
    Dim httpreq As New XMLHTTP60
    Dim strm As New Stream
    httpreq.Open "post", postUrl, False
    httpreq.setRequestHeader "contentType", "application/json; charset=UTF-8"
    httpreq.setRequestHeader "Cookie", cookie
    httpreq.send strPost
    If httpreq.Status <> 200 Then
        Exit Function
    End If
    Dim strRet As String
    strRet = httpreq.responseText
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Err.Clear
    End If
    Dim js As New JsonClass
    Set globeobj = js.GetJsonVal(strRet, "")
    Dim lpos As Long
    lpos = InStr(strRet, "SyncKey")
    If lpos > 0 Then
        Dim strkey As String
        Dim sUser As String
        strkey = Right(strRet, Len(strRet) - lpos)
        DecodeSyncKey (strkey)
        lpos = InStr(strkey, """User"": {")
        sUser = Right(strkey, Len(strkey) - lpos)
        
        Set SelfBean = New WXUser
        SelfBean.LoadDataBuyJson globeobj.User
'''        Sheet10.Cells(9, 12) = Trim(SelfBean.JsonObj.UserName)
'''        Sheet10.Cells(9, 13) = Trim(SelfBean.JsonObj.NickName)
    End If
    'Sheet10.Cells(3, 12) = strPost
    'Sheet10.Cells(4, 12) = strRet
    'Sheet10.Cells(2, 11) = postUrl
    'Sheet10.Cells(3, 10) = Cookie
    'Sheet10.Cells(3, 9) = SyncKey
    initSelf = True
End Function
Function getContactList() As Boolean
    On Error Resume Next
    Dim postUrl As String
''    strStamp = Sheet10.Cells(4, 8)
''    strHostFlg = Sheet10.Cells(5, 1)
    postUrl = "https://wx" & strHostFlg & ".qq.com/cgi-bin/mmwebwx-bin/webwxgetcontact"
    '{"BaseRequest":{"Uin":"UUUU","Sid":"SSSS","Skey":"KKKK","DeviceID":"e123456789012345"}}
    paramData = strPost
    Dim httpreq As New XMLHTTP60
    httpreq.Open "post", postUrl, False
    httpreq.setRequestHeader "contentType", "application/json; charset=UTF-8"
    httpreq.setRequestHeader "Cookie", cookie
    httpreq.send paramData
    Dim ContacterList As Dictionary
    If httpreq.Status <> 200 Then
        Exit Function
    End If
    Dim wxc As New WXContactor
    Set FriendList = wxc.LoadListByText(httpreq.responseText)
    getContactList = True
End Function

Sub DecodeSyncKey(str As String)
    Dim strkey As String
    Dim fullkey As String
    fullkey = Mid(str, InStr(str, "{") + 1, InStr(str, "]") - InStr(str, "{"))
    strkey = Mid(str, InStr(str, "["), InStr(str, "]") - InStr(str, "[") - 1)
    SyncKey = Replace(Replace(fullkey, Chr(13), ""), Chr(10), "")
    strkey = Replace(strkey, """", "")
    strkey = Replace(strkey, "Key:", "")
    strkey = Replace(strkey, "Val:", "_")
    strkey = Replace(strkey, "{", "")
    strkey = Replace(strkey, "}", "|")
    strkey = Replace(strkey, ",", "")
    strkey = Replace(strkey, "[", "")
    strkey = Replace(strkey, "]", "")
    strkey = Replace(strkey, Chr(13), "")
    strkey = Replace(strkey, Chr(10), "")
    strkey = Replace(strkey, " ", "")
    
    syncCheckKey = Left(strkey, Len(strkey) - 1)
    Sheet10.cells(1, 8) = SyncKey
    Sheet10.cells(2, 8) = syncCheckKey
End Sub

Sub decodeData_ticket(str As String)

End Sub
Function getTimeStap()
   Dim starTime As Date
   starttime = DateSerial(1970, 1, 1)
   getTimeStap = Trim(str(DateDiff("s", starttime, Now())))
End Function

Sub test()
    Dim obj As Image
    refreshCode obj
    'StartTimer2
    'getWx_Code
    
End Sub

''''''''''''''''Public Sub StartTimer()
''''''''''''''''    Dim RunWhen As Variant
''''''''''''''''    RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
''''''''''''''''        Schedule:=True
''''''''''''''''End Sub
''''''''''''''''
''''''''''''''''Public Sub StartTimer2()
''''''''''''''''    RunWhen = Now + TimeSerial(0, 0, 10)
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:="checkLogin", _
''''''''''''''''        Schedule:=True
''''''''''''''''End Sub
''''''''''''''''
''''''''''''''''Public Sub StartTimerGetMsg()
''''''''''''''''    RunWhen = Now + TimeSerial(0, 0, 60)
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:="getMsgData", _
''''''''''''''''        Schedule:=True
''''''''''''''''End Sub
''''''''''''''''
''''''''''''''''Public Sub StopTimer()
''''''''''''''''    On Error Resume Next
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
''''''''''''''''        Schedule:=False
''''''''''''''''End Sub
''''''''''''''''
''''''''''''''''Public Sub StopTimer2()
''''''''''''''''    On Error Resume Next
''''''''''''''''    RunWhen = Now + TimeSerial(0, 0, 10)
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:="checkLogin", _
''''''''''''''''        Schedule:=False
''''''''''''''''End Sub
''''''''''''''''
''''''''''''''''Public Sub StopTimerGetMsg()
''''''''''''''''    On Error Resume Next
''''''''''''''''    RunWhen = Now + TimeSerial(0, 0, 30)
''''''''''''''''    Application.OnTime EarliestTime:=RunWhen, Procedure:="getMsgData", _
''''''''''''''''        Schedule:=False
''''''''''''''''    testfalse = 1
''''''''''''''''End Sub

Public Sub refreshCode(pic As Image)
     getWx_Code pic
     'Sheet10.Cells(2, 1) = 0
     'StartTimer2
End Sub

Public Sub checkLogin()
''''''    'Sheet10.Cells(4, 10) = Now()
''''''    'UserForm1.Caption = Now()
''''''    Dim strCheckFlg As String
''''''    'strCheckFlg = Sheet10.Cells(2, 1).Text
''''''    If strCheckFlg = "1" Then Exit Sub
''''''    Dim strurl As String
''''''    If wait_login(strurl) Then
''''''        Sheet10.Cells(2, 1) = 1
''''''        UserForm1.StatusBar1.Panels(1).Text = "登陆成功！"
''''''         'StopTimer
''''''         'StopTimer2
''''''         'MsgBox "登陆成功！"
''''''         Init strurl
''''''         Exit Sub
''''''     End If
''''''     StartTimer2
End Sub

Sub init(strurl As String, sB As StatusBar)
    On Error Resume Next
    sB.Panels(2).Text = "获取Cookie。。。"
    If Not getCookie(strurl) Then
        MsgBox "无法获得基本信息！"
        Exit Sub
    End If
    sB.Panels(2).Text = "初始化用户信息。。。"
    If Not initSelf() Then
        MsgBox "无法初始化！"
        Exit Sub
    End If
    sB.Panels(2).Text = "获取联系人信息。。。"
    If Not getContactList() Then
        MsgBox "无法获取联系人列表！"
        Exit Sub
    End If
    ListContactor getMyAllContactors()
    'BindToDDL
    'getMsgData
End Sub

Sub ListContactor(ByRef AllContactors() As WXContactor)
    'Sheet11.Cells.Clear
    Dim struin As String
    Dim cnt As Long
    cnt = 0
    titleListed = False
    ReDim AllContactor(FriendList.Count)
    For Each obj In FriendList.Keys
        cnt = cnt + 1
        Dim wxc As WXContactor
        Set wxc = FriendList(obj)
        Dim col As Integer
        col = 0
        Set AllContactors(cnt) = New WXContactor
        AllContactors(cnt).LoadDataBuyJson wxc.JsonObj
        
''''        AllContactors(cnt).JsonObj = wxc.JsonObj.UserName
''''        Sheet11.Cells(cnt, 3) = wxc.JsonObj.NickName
''''        Sheet11.Cells(cnt, 4) = wxc.JsonObj.Signature
''''        Sheet11.Cells(cnt, 5) = wxc.JsonObj.ContactFlag
''''        Sheet11.Cells(cnt, 6) = wxc.JsonObj.VerifyFlag
''''        Sheet11.Cells(cnt, 7) = wxc.JsonObj.Province
''''        Sheet11.Cells(cnt, 8) = wxc.JsonObj.City
''''        Sheet11.Cells(cnt, 9) = wxc.JsonObj.HeadImgUrl
        titleListed = True
    Next
End Sub

Public Sub BindToDDL(dll As ComboBox)
    Dim struin As String
    Dim cnt As Long
    With dll
        .Clear
        Set .DataSource = FriendList
    End With
''''    With UserForm1.txt_ToUser
''''        .Clear
''''        .ColumnCount = 2
''''        .TextColumn = 2
''''        .BoundColumn = 1
''''        .Style = fmStyleDropDownList
''''        sheetranges = "WXContectors!B2:C" & Trim(str(FriendList.Count + 1))
''''        .RowSource = sheetranges
''''    End With
End Sub

Public Sub getNewMsg()
    On Error Resume Next
    UserForm1.Caption = Now()
    Dim msglist() As WXMsg
    Dim jc As New JsonClass
    msglist = getCurrentMsg()
    If DicFriends Is Nothing Then getFriendDic
    Sheet11.Range("$AH$1:$AL$1000").Clear
    For i = 1 To UBound(msglist)
        
        If DicFriends.Exists(msglist(i).FromUserName) Then
            Sheet11.cells(i, 34) = DicFriends(msglist(i).FromUserName)
        End If
''        If DicFriends.Exists(msglist(i).ToUserName) Then
''            Sheet11.Cells(i, 35) = DicFriends(msglist(i).ToUserName)
''        End If
        Sheet11.cells(i, 35) = msglist(i).Msg
        Sheet11.cells(i, 36) = msglist(i).FromUserName
        Sheet11.cells(i, 37) = msglist(i).MsgType
    Next
    UserForm1.txt_msgcnt = UBound(msglist)
    With UserForm1.txt_msglist
        .Clear
        Dim strRange As String
        strRange = "WXContectors!AH1:AL" & UBound(msglist)
        .RowSource = strRange
        .listStyle = fmListStylePlain
        .columnCount = 3
        .ColumnWidths = "60,380,60"
        .BoundColumn = 3
        '.ColumnWidths = 60
        .Selected(UBound(msglist) - 1) = True
    End With
   
    
    
End Sub

Function getCurrentMsg() As WXMsg()
    Dim ret() As WXMsg
    If synCheckGet() = -1 Then
        ReDim ret(0)
        getCurrentMsg = ret
        Exit Function
    End If
    getCurrentMsg = getCurrentMsgData()
End Function

Function synCheckGet() As Integer
    On Error Resume Next
    strStamp = getTimeStap()
    skey = Sheet10.cells(2, 12).Text
    wxsid = Sheet10.cells(2, 13).Text
    wxuin = Sheet10.cells(2, 14).Text
    pass_ticket = Sheet10.cells(2, 15).Text
    SyncKey = Sheet10.cells(1, 8)
    syncCheckKey = Sheet10.cells(2, 8)
    strHostFlg = Sheet10.cells(5, 1)
    Dim url As String
    url = "https://webpush.wx" & strHostFlg & ".qq.com/cgi-bin/mmwebwx-bin/synccheck?r=" & strStamp & "&skey=" & skey & "&sid=" & wxsid & "&uin=" & wxuin & "&deviceid=e" & Right("000000000000" & getTimeStap(), 15) & "&synckey=" + syncCheckKey + "&_=" + getTimeStap()
    Dim httpreq As New XMLHTTP60
    httpreq.Open "post", url, False
    httpreq.setRequestHeader "contentType", "text/html;charset=UTF-8"
    httpreq.setRequestHeader "Cookie", cookie
    httpreq.send
    If httpreq.Status <> 200 Then
        synCheckGet = -1
        Exit Function
    End If
    Dim js As New JsonClass
    Dim retobj As Object
    Dim strRet As String
    Dim strArr() As String
    strRet = httpreq.responseText
    strArr = Split(strRet, "=")
    Set retobj = js.GetJsonVal(strArr(1), "")
    If retobj.retcode <> "0" Then
'        synCheckGet = -1
'        MsgBox "心跳包错误！[" & strRet & "]" & Url
'        StopTimerGetMsg
'        refreshCode
        Exit Function
    End If
    If retobj.selector <> 2 And retobj.selector <> 6 Then
        synCheckGet = -1
        'MsgBox "无消息！"
        Exit Function
    End If
    synCheckGet = 1
End Function

Function getCurrentMsgData() As WXMsg()
    Dim postUrl As String
    Dim ret() As WXMsg
    strHostFlg = Sheet10.cells(5, 1)
    postUrl = "https://wx" & strHostFlg & ".qq.com/cgi-bin/mmwebwx-bin/webwxsync?sid=" & wxsid & "&skey=" & skey
    Dim strModel As String
    Dim strPost As String
    strStamp = getTimeStap()
    skey = Sheet10.cells(2, 12).Text
    wxsid = Sheet10.cells(2, 13).Text
    wxuin = Sheet10.cells(2, 14).Text
    pass_ticket = Sheet10.cells(2, 15).Text
    SyncKey = Sheet10.cells(1, 8)
    syncCheckKey = Sheet10.cells(2, 8)
    cookie = Sheet10.cells(1, 1).Text
    strModel = "{""BaseRequest"":{""Uin"":""UUUU"",""Sid"":""SSSS"",""Skey"":""KKKK"",""DeviceID"":""eDDDD""},""SyncKey"":{NNNN},""rr"":RRRR}"
    strPost = Replace(strModel, "UUUU", wxuin)
    strPost = Replace(strPost, "SSSS", wxsid)
    strPost = Replace(strPost, "KKKK", skey)
    strPost = Replace(strPost, "DDDD", Right("00000000000" & strStamp, 15))
    strPost = Replace(strPost, "NNNN", SyncKey)
    strPost = Replace(strPost, "RRRR", strStamp)
    Sheet10.cells(8, 11) = strPost
    'paramData = strPost
    Dim jc As New JsonClass
    Set obj = jc.GetJsonVal(strPost, "")
    If obj Is Nothing Then
        MsgBox "构建请求错误！" & strPost
       ' Exit Function
    End If
    Dim httpreq As New XMLHTTP60
    httpreq.Open "post", postUrl, False
    httpreq.setRequestHeader "contentType", "application/json; charset=UTF-8"
    httpreq.setRequestHeader "Cookie", cookie
    httpreq.send strPost
    If httpreq.Status <> 200 Then
        ReDim ret(0)
        getCurrentMsgData = ret
        Exit Function
    End If
    Dim Msg As New WXMsg
    
    Set obj = jc.GetJsonVal(httpreq.responseText, "")
    If obj.BaseResponse.ret <> 0 Then
        MsgBox obj.retcode
        Exit Function
    End If
    getCurrentMsgData = Msg.LoadMsgsByText(httpreq.responseText)
End Function

Function getFriendDic() As Dictionary
    Set DicFriends = New Dictionary
    For i = 2 To Sheet11.UsedRange.Rows.Count
        UserName = Trim(Sheet11.cells(i, 2).Text)
        NikeName = Trim(Sheet11.cells(i, 3).Text)
        If Not DicFriends.Exists(UserName) Then DicFriends.Add UserName, NikeName
    Next
End Function
Public Function SendMyMsg(ToUserName As String, Content As String) As Boolean
    Dim strMyInfo As String
    Dim jc As New JsonClass
    strMyInfo = Sheet10.cells(9, 12).Text
    SendMyMsg = SendMsg(strMyInfo, ToUserName, Content)
End Function
Function SendMsg(FromUserName As String, ToUserName As String, Content As String) As Boolean
    On Error Resume Next
    If Len(Trim(FromUserName)) = 0 Or Len(Trim(ToUserName)) = 0 Then Exit Function
    If Len(Trim(Content)) = 0 Then
        MsgBox " 内容不能为空!"
        Exit Function
    End If
    Dim postUrl As String
    Dim strModel As String
    Dim strPost As String
    strStamp = getTimeStap()
    skey = Sheet10.cells(2, 12).Text
    wxsid = Sheet10.cells(2, 13).Text
    wxuin = Sheet10.cells(2, 14).Text
    pass_ticket = Sheet10.cells(2, 15).Text
    SyncKey = Sheet10.cells(1, 8).Text
    syncCheckKey = Sheet10.cells(2, 8).Text
    cookie = Sheet10.cells(1, 1).Text
    strHostFlg = Sheet10.cells(5, 1)
    postUrl = "https://wx" & strHostFlg & ".qq.com/cgi-bin/mmwebwx-bin/webwxsendmsg?pass_ticket=" & pass_ticket
    strModel = "{""BaseRequest"":{""Uin"":""UUUU"",""Sid"":""SSSS"",""Skey"":""KKKK"",""DeviceID"":""eDDDD""}," _
    & """Msg"":{""Type"":1,""Content"":""{Content}"",""FromUserName"":""{FromUserName}"",""ToUserName"":""{toUserName}""," _
    & """LocalID"":""TTTT"",""ClientMsgId"":""TTTT""},""Scene"":0,""rr"":TTTT}"
    strPost = Replace(strModel, "UUUU", wxuin)
    strPost = Replace(strPost, "SSSS", wxsid)
    strPost = Replace(strPost, "KKKK", skey)
    strPost = Replace(strPost, "DDDD", Right("111111111111111" & strStamp, 15))
    strPost = Replace(strPost, "TTTT", strStamp)
    strPost = Replace(strPost, "{FromUserName}", FromUserName)
    strPost = Replace(strPost, "{Content}", Content)
    strPost = Replace(strPost, "{toUserName}", ToUserName)
    Sheet10.cells(11, 12) = strPost
    tempData = "refreshTimes=5; login_frequency=2; " & cookie
    Dim jc As New JsonClass
    Set testobj = jc.GetJsonVal(strPost, "")
    Dim httpreq As New XMLHTTP60
    httpreq.Open "post", postUrl, False
    httpreq.setRequestHeader "contentType", "application/json; charset=UTF-8"
    httpreq.setRequestHeader "Cookie", tempData
    httpreq.send strPost
    If httpreq.Status <> 200 Then
        ReDim ret(0)
        Exit Function
    End If
    Dim strRet As String
    strRet = httpreq.responseText
    Set obj = jc.GetJsonVal(strRet, "")
    If obj.BaseResponse.ret = 0 Then
        SendMsg = True
    End If
End Function



