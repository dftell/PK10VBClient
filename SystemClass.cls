VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SysParams As Dictionary 'config_xx_xx
Public currExpect As String
Public lastExpect As String
Public LastHistoryExpect As String
Public FirstValideHistoryExpect As String
Public LastInValideHistoryExpect As String
Dim configDoc As DOMDocument60
Public WxMyInfo As WXUser
Public WxSendUser As WXContactor
Public WxSpeaker As WXContactor
Public NoHtmlGetNewData As Boolean
Public InterStatus As Boolean
Public ExchangeStatus As Boolean
Public WXLogined As Boolean
Public LogObj As New LogClass
Public Ip As Integer
Dim t_newExpectData() As ExpectData
Public b_AllowExchange As Boolean
Public InstsList As Dictionary
Public ExchangeSwitched As Boolean

Public Property Get AllowExchange() As Boolean
    If SysParams.Count > 0 Then b_AllowExchange = IIf(SysParams("System")("AllowExchange") = "1", True, False)
    AllowExchange = b_AllowExchange
End Property


Public Property Let AllowExchange(val As Boolean)
    b_AllowExchange = val
End Property

Public Property Get TotalCnt() As Integer
    If SysParams.Count > 0 Then TotalCnt = CInt(SysParams("System")("TotalCnt"))
End Property

Public Property Get AllowHedge() As Boolean
    If SysParams.Count > 0 Then AllowHedge = CBool(SysParams("System")("AllowHedge"))
End Property

Public Property Get JoinHedge() As Boolean
    If SysParams.Count > 0 Then JoinHedge = CBool(SysParams("System")("JoinHedge"))
End Property

Public Property Get HedgeTimes() As Long
    If SysParams.Count > 0 Then HedgeTimes = CLng(SysParams("System")("HedgeTimes"))
End Property

Public Property Get IsClient() As Boolean
    If SysParams.Count > 0 Then IsClient = IIf(SysParams("System")("IsClient") = "1", True, False)
End Property

Public Property Get ClientUserName() As String
    If SysParams.Count > 0 Then ClientUserName = SysParams("System")("ClientUsername")
End Property

Public Property Get ClientPassword() As String
    If SysParams.Count > 0 Then ClientPassword = SysParams("System")("ClientPassword")
End Property

Public Property Get LoginUrlModel() As String
    If SysParams.Count > 0 Then LoginUrlModel = SysParams("System")("LoginUrlModel")
End Property

Public Property Get LoginDefaultHost() As String
    If SysParams.Count > 0 Then LoginDefaultHost = SysParams("System")("LoginDefaultHost")
End Property

Public Property Get LoginHostList() As String
    If SysParams.Count > 0 Then LoginHostList = SysParams("System")("LoginHostList")
End Property

Public Property Get LoginInstFillOrEnCode() As Integer
    If SysParams.Count > 0 Then LoginInstFillOrEnCode = CInt(SysParams("System")("LoginInstFillOrEnCode"))
End Property


Public Property Get LoginInFrame() As Integer
    If SysParams.Count > 0 Then LoginInFrame = CInt(SysParams("System")("LoginInFrame"))
End Property

Public Property Get MinChips() As Integer
    If SysParams.Count > 0 Then MinChips = CInt(SysParams("System")("MinChips"))
End Property


Public Property Get SingleColMinTimes() As Integer
    If SysParams.Count > 0 Then SingleColMinTimes = CInt(SysParams("System")("SingleColMinTimes"))
End Property


Public Property Get NewestExpectData() As ExpectData
    Set NewestExpectData = t_newExpectData(1)
End Property

Public Property Get CurrExpectData() As ExpectData()
    CurrExpectData = t_newExpectData
End Property

Public Sub SetCurrExpectData(val() As ExpectData)
    ReDim t_newExpectData(0)
    If UBound(val) = 0 Then
        Exit Sub
    End If
    t_newExpectData = val
End Sub

Public Property Get StartCols() As Integer
    If SysParams.Count > 0 Then StartCols = CLng(SysParams("System")("StartCol"))
End Property

Public Property Get Odds() As Currency
    If SysParams.Count > 0 Then Odds = CCur(SysParams("System")("Odds"))
End Property

Public Property Get InterVal() As Integer
    If SysParams.Count > 0 Then InterVal = CLng(SysParams("System")("InterVal"))
End Property

Public Property Get BackColor() As Long
    BackColor = RGB(CInt(SysParams("System")("BackColor_R")), CInt(SysParams("System")("BackColor_G")), CInt(SysParams("System")("BackColor_B")))
End Property

Public Property Get HistoryFromPage() As Long
    If SysParams.Count > 0 Then HistoryFromPage = CLng(SysParams("Research")("FromPage"))
End Property

Public Property Get NewestHistoryExpect() As Long
    If SysParams.Count > 0 Then NewestHistoryExpect = CLng(SysParams("Research")("NewestHistoryExpect"))
End Property

Public Property Get MutliColMinTimes() As Integer
    If SysParams.Count > 0 Then MutliColMinTimes = CInt(SysParams("System")("MutliColMinTimes"))
End Property

Public Property Get SingleCarRepeatCnt() As Integer
    If SysParams.Count > 0 Then SingleCarRepeatCnt = CInt(SysParams("Research")("SingleCarRepeatCnt"))
End Property

Public Property Get RepeatCheckCnt() As String
    If SysParams.Count > 0 Then RepeatCheckCnt = SysParams("Research")("RepeatCheckCnt")
End Property

Public Property Get ResearchStartCol() As Integer
    If SysParams.Count > 0 Then ResearchStartCol = CInt(SysParams("Research")("StartCol"))
End Property
Public Property Get ValidOldestHistoryExpect() As String
    If SysParams.Count > 0 Then ValidOldestHistoryExpect = SysParams("Research")("ValidOldestHistoryExpect")
End Property

Public Property Get AssetInitCash() As Long
    If SysParams.Count > 0 Then AssetInitCash = SysParams("Asset")("InitCash")
End Property

Public Property Get AssetCosted() As Long
    If SysParams.Count > 0 Then AssetCosted = SysParams("Asset")("Costed")
End Property

Public Property Get AssetGained() As Single
    If SysParams.Count > 0 Then AssetGained = SysParams("Asset")("Gained")
End Property

Public Property Let AssetCosted(val As Long)
    If SysParams.Count > 0 Then SysParams("Asset")("Costed") = val
End Property

Public Property Let AssetGained(val As Single)
    If SysParams.Count > 0 Then SysParams("Asset")("Gained") = val
End Property

Public Property Get AssetAChanceMaxRate() As Single
    If SysParams.Count > 0 Then AssetAChanceMaxRate = SysParams("Asset")("AChanceMaxRate")
End Property

Public Property Get AssetTotalMaxRate() As Single
    If SysParams.Count > 0 Then AssetTotalMaxRate = SysParams("Asset")("TotalMaxRate")
End Property

Public Property Get AssetTotal() As Single
    If SysParams.Count > 0 Then AssetTotal = SysParams("Asset")("TotalCash")
End Property

Public Property Let AssetTotal(val As Single)
    If SysParams.Count > 0 Then SysParams("Asset")("TotalCash") = val
End Property

Public Property Get MinTimeForChance(Times As Integer) As Integer
    Dim strTimes As String
    If Times < 1 Or Times > 8 Then
        'MsgBox times
        Dim test As String
        test = ""
    End If
    strTimes = "MinTimesFor" & Trim(CStr(Times))
    If SysParams.Count > 0 Then
        MinTimeForChance = CInt(SysParams("Exchange")(strTimes)) 'MinTimesFor2
    End If
End Property

Public Property Get UnitChipArray(Cols As Integer) As String()
    If SysParams.Count > 0 Then UnitChipArray = Split(SysParams("Exchange")("Serial" & CStr(Cols)), ",")
End Property

Public Property Get SerTotal(Cols As Integer) As Integer
    On Error Resume Next
    If SysParams.Count > 0 Then
        Dim edic As Dictionary
        Set edic = SysParams("Exchange")
        If Not edic.Exists("SerTotal" & Cols) Then
            SerTotal = 1
            Exit Property
        End If
        SerTotal = CInt(SysParams("Exchange")("SerTotal" & CStr(Cols)))
    End If
End Property

Public Sub InitSystemParams()

End Sub

Sub getConfigInfo()
    On Error Resume Next
    Dim configs As IXMLDOMNodeList
    Set configs = configDoc.SelectNodes("root/configs/config")
    Set SysParams = Nothing
    Set SysParams = New Dictionary
    For i = 0 To configs.Length - 1
        Dim node As IXMLDOMNode
        Set node = configs(i)
        Dim typename As String
        typename = node.Attributes(0).NodeValue
        Dim configtypeDir As New Dictionary
        SysParams.Add typename, New Dictionary
        Set configtypeDir = SysParams(typename)
        Dim configitems As IXMLDOMNodeList
        Set configitems = node.SelectNodes("./item")
        For j = 0 To configitems.Length - 1
            Dim Name As String
            Dim Value As String
            Name = configitems(j).SelectSingleNode("@key").NodeValue
            Value = configitems(j).SelectSingleNode("@value").NodeValue
            SysParams(typename).Add Name, Value
        Next
        Set configtypeDir = Nothing
    Next
End Sub

Sub setConfigInfo()
    On Error Resume Next
    Dim configs As IXMLDOMNodeList
    Set configs = configDoc.SelectNodes("root/configs/config")
    'Set SysParams = New Dictionary
    For i = 0 To configs.Length - 1
        Dim node As IXMLDOMNode
        Set node = configs(i)
        Dim typename As String
        typename = node.Attributes(0).NodeValue
        Dim configtypeDir As New Dictionary
        Set configtypeDir = SysParams(typename)
        'Dim configitems As IXMLDOMNodeList
        'Set configitems = node.SelectNodes("item")
        Dim strname As Variant
        For Each strname In configtypeDir.Keys
           Dim iNode As IXMLDOMNode
           Set iNode = node.SelectSingleNode(Replace("item[@key='XX']", "XX", strname))
           iNode.SelectSingleNode("@value").NodeValue = configtypeDir(strname)
        Next
        configDoc.Save App.Path & "\config.xml"
''        For j = 0 To configitems.Length - 1
''            Dim Name As String
''            Dim Value As String
''            Name = configitems(j).SelectSingleNode("@key").NodeValue
''            configitems(j).SelectSingleNode("@value").NodeValue
''            configtypeDir.Add Name, Value
''        Next
    Next
End Sub


Sub getXml()
    'Dim xmldoc As New IXMLDOMComment
    Dim xmldoc As New DOMDocument60
    If xmldoc.Load(App.Path & "\config.xml") Then
    If Err.Number <> 0 Then Me.LogObj.Log Err.Description
        'MsgBox Err.Description
    Set configDoc = xmldoc
    Else
        'MsgBox "can't find the config file!"
        Me.LogObj.Log "can't find the config file!"
    End If
End Sub

Private Sub Class_Initialize()
    getXml
    getConfigInfo
End Sub

Public Sub FillControl(frm As Form)
    On Error Resume Next
    Dim strtype As Variant
    For Each strtype In SysParams.Keys
        Dim strItem As Variant
        Dim itemDir As Dictionary
        Set itemDir = SysParams(strtype)
        For Each strItem In itemDir.Keys
            Dim configitem As New ConfigClass
            Set configitem.theForm = frm
            configitem.ConfigType = strtype
            configitem.Name = strItem
            configitem.Value = itemDir(strItem)
            Dim txtctrl As TextBox
            Set txtctrl = configitem.Contrl
            If txtctrl Is Nothing Then
            Else
                txtctrl.Text = configitem.Value
            End If
        Next
    Next
End Sub

Public Sub SaveConfig(frm As Form)
    On Error Resume Next
    Dim strtype As Variant
    For Each strtype In SysParams.Keys
        Dim strItem As Variant
        Dim itemDir As Dictionary
        Set itemDir = SysParams(strtype)
        For Each strItem In itemDir.Keys
            Dim configitem As New ConfigClass
            configitem.Name = strItem
            configitem.ConfigType = strtype
            Set configitem.theForm = frm
            Dim txtctrl As TextBox
            Set txtctrl = configitem.Contrl
            If txtctrl Is Nothing Then
            Else
                 itemDir(strItem) = txtctrl.Text
            End If
        Next
    Next
    Me.setConfigInfo
End Sub

Sub test()
    Dim fo As New FileSystemObject
    Dim tobj
    
    
    
End Sub
