VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Public gobj As SystemClass
Public strurl As String 'URL
Public iInterMinutes As Double '触发间隙
Public RunWhen As Double
Public cRunIntervalSeconds As Long ' 120 ' two minutes
Public iMinCheckTimes  As Integer '最少未出现次数
Public iColShift As Integer
Public iSelectColumns  As Integer '
Public bAutoSchedule As Boolean '是否自动刷新
Public iOneCodeMinCheckTimes As Integer
Const MaxPos = 10
Const cRunWhat = "RefreshData"  ' the name of the procedure to run
Public strAllSingleChances As String
Public currExpect As String
Public Assetobj As New AssetClass
Public chanceDic As New Dictionary
Public CurrTimeHaveFClose As Boolean
Public DicFClose As Dictionary
Public DicAllUseData As Dictionary
Public DicAllHistoryData As Dictionary
Public frm As Form
Public NewestExpect As String
'Public AccessHistoryPages As Long
Public ValidOldestHistoryExpect As String
Public CurrHistoryExpect As String
Public NewestHistoryExpect As String

Dim NewestData() As ExpectData
Dim NewestHistory() As ExpectData
Dim t_fNewestData As Boolean
Public ExpectChancesDic As New Dictionary
Public ExpectInstructStatusDic As New Dictionary
Const ISingleMin = 35
Const ISingleMax = 60
Public FirstRepeatCnt As Integer

Public Property Get CurrDataCount() As Integer
    If UBound(NewestData) <= 0 Then
        CurrDataCount = 0
    Else
        CurrDataCount = UBound(NewestData)
    End If
End Property
Public Property Get fNewestData() As Boolean
    fNewestData = t_fNewestData
End Property

Public Property Let fNewestData(val As Boolean)
    t_fNewestData = val
End Property

Public Sub StartTimer()
    RunWhen = Now + TimeSerial(0, 0, cRunIntervalSeconds)
    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
        Schedule:=bAutoSchedule
End Sub
Public Sub getParams()
    On Error Resume Next
    strurl = Sheet2.cells(1, "b").Text
    cRunIntervalSeconds = Sheet2.cells(2, "b").Text * 60
    iMinCheckTimes = Sheet2.cells(3, "b").Text
    iColShift = Sheet2.cells(4, "b").Text
    iSelectColumns = Sheet2.cells(5, "b").Text
    bAutoSchedule = IIf(Trim(Sheet2.cells(6, "b").Text) = "1", True, False)
    iOneCodeMinCheckTimes = Sheet2.cells(7, "b").Text
    strAllSingleChances = Sheet1.cells(3, 11).Text
    currExpect = Sheet1.cells(1, 2).Text
    SingleMinTimes = Sheet1.cells(1, 2).Text
    LoadAssetInfo
    InitChanceList
End Sub

Public Sub InitParams(sysconfig As SystemClass)
    On Error Resume Next
    strurl = sysconfig.SysParams("System")("URL")
    iMinCheckTimes = sysconfig.SysParams("System")("MutliColMinTimes")
    'iSelectColumns = sysconfig.SysParams("System")("StartCol")
    iColShift = sysconfig.SysParams("System")("StartCol")
    iOneCodeMinCheckTimes = sysconfig.SingleColMinTimes
    currExpect = sysconfig.currExpect
    Me.CurrHistoryExpect = sysconfig.NewestHistoryExpect
    ValidOldestHistoryExpect = gobj.ValidOldestHistoryExpect
    
''''    strurl = .txt_config_System_NewestDataURL.Text
''''    cRunIntervalSeconds = Sheet2.Cells(2, "b").Text * 60
''''    iMinCheckTimes = Sheet2.Cells(3, "b").Text
''''    iColShift = Sheet2.Cells(4, "b").Text
''''    iSelectColumns = Sheet2.Cells(5, "b").Text
''''    bAutoSchedule = IIf(Trim(Sheet2.Cells(6, "b").Text) = "1", True, False)
''''    iOneCodeMinCheckTimes = Sheet2.Cells(7, "b").Text
''''    strAllSingleChances = Sheet1.Cells(3, 11).Text
''''    CurrExpect = Sheet1.Cells(1, 2).Text
''''    SingleMinTimes = Sheet1.Cells(1, 2).Text
''''    LoadAssetInfo
''''    InitChanceList
   
End Sub

Public Sub LoadAssetInfo()
    Set Assetobj = New AssetClass
    With gobj
        Assetobj.InitCash = .AssetInitCash
        Assetobj.CurrCash = .AssetCosted
        Assetobj.Gained = .AssetGained
        Assetobj.OneChanceMaxLossRateToInit = .AssetAChanceMaxRate
        Assetobj.OneChanceMaxLossRateToTotal = .AssetAChanceMaxRate
        Assetobj.MaxLossAbsToTotal = .AssetTotalMaxRate
        Assetobj.TotalAsset = .AssetTotal
    End With
''    With Sheet1
''        Assetobj.InitCash = CLng(Sheet2.Cells(1, 7).Text)
''        Assetobj.CurrCash = CLng(Sheet2.Cells(2, 7).Text)
''        Assetobj.Gained = CLng(Sheet2.Cells(3, 7).Text)
''        Assetobj.TotalAsset = CLng(Sheet2.Cells(4, 7).Text)
''        Assetobj.OneChanceMaxLossRateToInit = Sheet2.Cells(5, 7).Value
''        Assetobj.OneChanceMaxLossRateToTotal = Sheet2.Cells(6, 7).Value
''        Assetobj.MaxLossAbsToTotal = Sheet2.Cells(7, 7).Value
''        Assetobj.MaxGainAbsToInit = Sheet2.Cells(8, 7).Value
''    End With
End Sub

Sub UpdateAsset()
    With gobj
        .AssetCosted = Assetobj.CurrCash
        .AssetGained = Assetobj.Gained
        .AssetTotal = Assetobj.TotalAsset
    End With
''''' With Sheet1
'''''    Sheet2.Cells(2, 7) = Assetobj.CurrCash
'''''    Sheet2.Cells(3, 7) = Assetobj.Gained
'''''    Sheet2.Cells(4, 7) = Assetobj.TotalAsset
'''''    End With
End Sub

Public Sub StopTimer()
    On Error Resume Next
    Application.OnTime EarliestTime:=RunWhen, Procedure:=cRunWhat, _
        Schedule:=False
End Sub

Public Sub refreshData()
    getParams
    If bAutoSchedule Then
'        StopTimer
'        StartTimer
    End If
    'Process
    
End Sub

 Sub run_Click()
    MsgBox ""
End Sub

Function ReadData(ByRef ret() As ExpectData)
    On Error Resume Next
    Dim dt As DataTable
    Dim ed As New ExpectData
    Set dt = ed.getDayData(DateAdd("D", -1, Date), currExpect)
    ed.TableToArray dt, ret
    Exit Function
    Dim cnt As Long
    With Sheet3
        'If .UsedRange.Rows.Count > 0 Then ReDim ret(.UsedRange.Rows.Count - 1)
        Dim i As Long
        For i = .UsedRange.Rows.Count To 2 Step -1
            Dim rec As New ExpectData
            rec.OpenTime = .cells(i, 3)
            If Year(rec.OpenTime) <> Year(Now()) Or Month(rec.OpenTime) <> Month(Now()) Or day(rec.OpenTime) <> day(Now()) Then
                ReadData = ret
                Exit Function
            End If
            cnt = cnt + 1
            ReDim Preserve ret(cnt)
            Set ret(cnt) = New ExpectData
            ret(cnt).Expect = .cells(i, 1)
            'ret(cnt).expect = .Cells(i, 1)
            ret(cnt).OpenCode = .cells(i, 2)
            ret(cnt).OpenTime = .cells(i, 3)
            Set rec = Nothing
        Next
    End With
    ReadData = ret
End Function

Function ReadHistoryData(ByRef ret() As ExpectData)
    On Error Resume Next
    Dim ed As New ExpectData
    ed.TableToArray ed.GetHistoryData(Me.ValidOldestHistoryExpect, 100), ret
    Exit Function
    
    
    Dim cnt As Long
    With Sheet19
        'If .UsedRange.Rows.Count > 0 Then ReDim ret(.UsedRange.Rows.Count - 1)
        Dim i As Long
        For i = .UsedRange.Rows.Count To 2 Step -1
            Dim rec As New ExpectData
            rec.OpenTime = .cells(i, 3)
''            If Year(rec.opentime) <> Year(Now()) Or Month(rec.opentime) <> Month(Now()) Or Day(rec.opentime) <> Day(Now()) Then
''                ReadData = ret
''                Exit Function
''            End If
            cnt = cnt + 1
            ReDim Preserve ret(cnt)
            Set ret(cnt) = New ExpectData
            ret(cnt).Expect = .cells(i, 1)
            'ret(cnt).expect = .Cells(i, 1)
            ret(cnt).OpenCode = .cells(i, 2)
            ret(cnt).OpenTime = .cells(i, 3)
            Set rec = Nothing
        Next
    End With
    ReadHistoryData = ret
End Function

Function getCurrDataDic(data() As ExpectData)
    Dim ret As New Dictionary
    Dim i As Long
    For i = 1 To UBound(data)
        Dim rec As New ExpectData
        Set rec = data(i)
        If Len(Trim(rec.OpenCode)) > 0 And Not ret.Exists(rec.Expect) Then
            ret.Add rec.Expect, rec
        End If
    Next
    Set getCurrDataDic = ret
End Function

Sub WriteData(datas() As ExpectData, currdaydata As Dictionary)
    Dim addNewData() As ExpectData
    Dim i As Long
    Dim ai As Long
    For i = UBound(datas) To 1 Step -1
        Dim cnt As Long
        Dim rec As ExpectData
        Set rec = datas(i)
'        'With Sheet3
'            cnt = .UsedRange.Rows.Count
            If Not currdaydata.Exists(rec.Expect) And day(rec.OpenTime) = day(Now()) And Month(rec.OpenTime) = Month(rec.OpenTime) Then
                ai = ai + 1
                ReDim Preserve addNewData(ai)
                Set addNewData(ai) = New ExpectData
                With addNewData(ai)
                    .Expect = rec.Expect
                    .OpenTime = rec.OpenTime
                    .OpenCode = rec.OpenCode
                End With
            End If
        'End With
    Next
    Dim ed As New ExpectData
    Call ed.SaveNewData(addNewData)
    
End Sub

Function getAllDataDic(data() As ExpectData, ret As Dictionary)
    'Dim ret As New Dictionary
    Dim i As Long
    For i = UBound(data) To 1 Step -1
        Dim rec As New ExpectData
        Set rec = data(i)
        If Len(Trim(rec.OpenCode)) > 0 And Not ret.Exists(rec.Expect) Then
            ret.Add rec.Expect, rec
        Else
            'Set getAllDataDic = ret '遇到相同的就退出，节约时间
            'Exit Function
        End If
    Next
    Set getAllDataDic = ret
End Function
Sub WriteAllData(datas() As ExpectData, ByRef currdaydata As Dictionary)
    Dim i As Long
    Dim AddData() As ExpectData
    Dim AddI As Long
    For i = UBound(datas) To 1 Step -1
        Dim cnt As Long
        Dim rec As ExpectData
        Set rec = datas(i)
        If Not currdaydata.Exists(rec.Expect) Then
            AddI = AddI + 1
            ReDim Preserve AddData(AddI)
            Set AddData(AddI) = New ExpectData
            AddData(AddI).Expect = rec.Expect
            AddData(AddI).OpenTime = rec.OpenTime
            AddData(AddI).OpenCode = rec.OpenCode
            currdaydata.Add rec.Expect, AddData(AddI)
        End If
    Next
    Dim ed As New ExpectData
    ed.SaveHistoryData AddData
End Sub
Public Function RefreshNewestData(datas() As ExpectData, Optional NoHtml As Boolean) As Boolean
    On Error Resume Next
    'Dim datas() As ExpectData
    Dim currdaydata() As ExpectData
    Dim currdaydic As Dictionary
    ReDim currdaydata(0)
    If Not NoHtml Then
        'gobj.LogObj.Log "读取新数据"
        datas = ReReadHtml()
        If Me.NewestExpect = Me.currExpect Then
            Exit Function '未有新的数据，不刷新
        End If
        currExpect = NewestExpect
        gobj.lastExpect = CStr(CLng(NewestExpect) + 1)
        gobj.LogObj.OutExpectNo gobj.lastExpect
    End If
    ReadData currdaydata
    Dim dt As DataTable
    Dim ed As New ExpectData
    'test Set dt = ed.ArrayToTable(currdaydata)
    Set currdaydic = getCurrDataDic(currdaydata)
    If Not NoHtml Then
        WriteData datas, currdaydic
    End If
    ed.TableToArray ed.getDayData(DateAdd("D", -1, Date), currExpect), datas '修改为前一天
    If IsEmpty(datas) Or UBound(datas) = 0 Then Exit Function
    If Err.Number <> 0 Then Exit Function
    NewestData = datas
    gobj.SetCurrExpectData datas
    Set Me.DicAllUseData = Nothing
    Set Me.DicAllUseData = New Dictionary
    Dim c As Variant
    For Each c In Me.DicAllHistoryData.Keys
        Me.DicAllUseData.Add c, New ExpectData
        Me.DicAllUseData(c).Expect = Me.DicAllHistoryData(c).Expect
        Me.DicAllUseData(c).OpenTime = Me.DicAllHistoryData(c).OpenTime
        Me.DicAllUseData(c).OpenCode = Me.DicAllHistoryData(c).OpenCode
    Next
    For i = UBound(datas) To 1 Step -1
        If Not DicAllUseData.Exists(datas(i).Expect) Then
            Dim expectNo As String
            expectNo = datas(i).Expect
            DicAllUseData.Add expectNo, New ExpectData
            DicAllUseData(expectNo).Expect = datas(i).Expect
            DicAllUseData(expectNo).OpenTime = datas(i).OpenTime
            DicAllUseData(expectNo).OpenCode = datas(i).OpenCode
        End If
    Next
    gobj.LogObj.Log "共计期数：" & CStr(DicAllUseData.Count)
    RefreshNewestData = True
End Function
Public Sub Process(allgrid As MSFlexGrid, colgrid As MSFlexGrid, tbarr As String)
    On Error Resume Next
    Dim datas() As ExpectData
    'RefreshNewestData datas
    Dim ed As New ExpectData
    'ed.TableToArray ed.getDayData(Date), datas
    If IsArray(NewestData) And UBound(NewestData) <= 0 Then Exit Sub
    datas = NewestData
    Dim p As Integer
    Dim i As Integer
    
    Dim BuyCarArr() As String
    Dim BuyCnt As Integer
    For p = iColShift + 1 To MaxPos
       Dim maxTimes As Integer
       Dim maxCarNo As String
       maxTimes = 0
       For i = 1 To 10
         currTime = getInColumnNoTimes(datas, i, p)
         If currTime > maxTimes Then
             maxTimes = currTime
             maxCarNo = Replace(str(i), "10", "0")
         End If
       Next
       Dim strCarNo As String
       strCarNo = Replace(Replace(str(p), "10", "0") & "/" & maxCarNo, " ", "")
       'Sheet1.Cells(4, 9 + p) = Replace(Replace(str(p), "10", "0") & "/" & maxCarNo & ":" & str(maxTimes), " ", "")
       Dim ctrl As TextBox
       Set ctrl = CallByName(colgrid.Parent, tbarr, VbGet, p - 1)
       ctrl.Text = Replace(Replace(str(p), "10", "0") & "/" & maxCarNo & ":" & str(maxTimes), " ", "")
       If maxTimes > iOneCodeMinCheckTimes Then
           'Sheet1.Cells(4, 9 + p).Interior.Color = RGB(255, 255, 0)
           ctrl.BackColor = vbYellow
           BuyCnt = BuyCnt + 1
           iMoney = maxTimes - iOneCodeMinCheckTimes + 1
           strbuy = strCarNo '& "/" & Sheet5.Cells(iMoney, "B")
           ReDim Preserve BuyCarArr(BuyCnt)
           BuyCarArr(BuyCnt) = strbuy
       Else
          'Sheet1.Cells(4, 9 + p).Interior.Color = RGB(255, 255, 255)
          ctrl.BackColor = vbWhite
       End If
    Next
    Dim strAllBuyInfo As String
   'Dim d As New DataObject
   strAllBuyInfo = Join(BuyCarArr, " ")
'''''   If Len(strAllBuyInfo) > 0 Then
'''''      d.SetText strAllBuyInfo
'''''      d.PutInClipboard
'''''      Set d = Nothing
'''''    End If
   'Sheet1.Cells(3, 11) = strAllBuyInfo 'all single chance
   '单列信息
   ProcessColNoCars datas, colgrid
    '综合信息
   For p = 2 To 8
        'iSelectColumns = p
        Dim combs() As Variant
        CombResult MaxPos - iColShift, p, combs
        Dim lastcombs() As Variant
        UpdateComb combs, iColShift, lastcombs
        'WriteComb lastcombs
'''''        For i = 0 To 31
'''''            Dim rowid As Integer
'''''            rowid = 6 + i
'''''            strRange = Replace("B#:S#", "#", rowid)
'''''            'Sheet1.Range(strRange).Borders.LineStyle = xlContinuous
'''''            If (i Mod 2) = 1 Then
'''''                'Sheet1.Range(strRange).Interior.Color = RGB(189, 215, 238)
'''''            End If
'''''            'If i = 10 Or i = 7 Or i = 4 Then Sheet1.Range(strRange).Interior.Color = RGB(255, 255, 0)
'''''        Next
        Dim maxCheckLong As Integer
        maxCheckLong = IIf(100 > UBound(datas), UBound(datas), 100)
        For i = iMinCheckTimes To maxCheckLong
            Dim testRet() As String
            Dim result() As String
            SearchCombRecords lastcombs, datas, i, p, testRet
            If UBound(testRet) = 0 Then
                result = testRet
                ReDim testRet(0)
                GoTo endpress
            End If
            'Sheet1.Cells(6 + i - iMinCheckTimes, p - 1) = Join(testRet, " ")
            If p < 5 Then
                Dim testRetBig() As String
                ReDim testRetBig(0)
                Dim k As Integer
                Dim kcnt As Integer
                kcnt = 0
                k = 0
                For k = 1 To UBound(testRet)
                    Dim cars As String
                    cars = Split(testRet(k), "/")(1)
                    If (p = 2 And Len(cars) > 2) Or (p > 2 And Len(cars) > 1) Then
                        kcnt = kcnt + 1
                        ReDim Preserve testRetBig(kcnt)
                        testRetBig(kcnt) = testRet(k)
                    End If
                Next
                If UBound(testRetBig) > 0 Then
                    If p = 2 Then
                        allgrid.TextMatrix(i - iMinCheckTimes + 1, 5) = Trim(allgrid.TextMatrix(i - iMinCheckTimes + 1, 5) + " " + Trim(Join(testRetBig, " ")))
                    Else
                        allgrid.TextMatrix(i - iMinCheckTimes + 1, p - 2 + 1 + p) = Trim(allgrid.TextMatrix(i - iMinCheckTimes + 1, p - 2 + 1 + p) + " " + Trim(Join(testRetBig, " ")))
                    End If
                End If
            End If
            If p > 2 Then allgrid.TextMatrix(i - iMinCheckTimes + 1, p - 2 + 1) = Trim(allgrid.TextMatrix(i - iMinCheckTimes + 1, p - 2 + 1) + " " + Trim(Join(testRet, " ")))

            result = testRet
            ReDim testRet(0)
        Next
endpress:
    Next
    'Research
    'GetHistoryData
    Dim dg As New DataGridClass
    Set dg.grid = allgrid
    'If Hour(Now()) < 10 Then dg.AdjustColWidth allgrid.Parent, allgrid
    allgrid.FixedRows = 1
    allgrid.AllowUserResizing = flexResizeColumns
    Set dg.grid = colgrid
    dg.AdjustColWidth colgrid.Parent, colgrid
    colgrid.FixedRows = 1
End Sub

Function getInColumnNoTimes(data() As ExpectData, carno As Integer, colno As Integer) As Long
    On Error Resume Next
    Dim i As Long
    For i = 1 To UBound(data)
        Dim code As String
        Dim strCar As String
        code = Split(data(i).OpenCode, ",")(colno - 1)
        If Err.Number <> 0 Then
            test = 1
        End If
        strCar = Right("00" & carno, 2)
        If carno = 5 And colno = 7 And i = 23 Then
            Dim testflg As Integer
            testflg = 1
        End If
        If code = strCar Then
            getInColumnNoTimes = i - 1
            Exit Function
        End If
    Next
    getInColumnNoTimes = UBound(data)
End Function

Sub ProcessColNoCars(data() As ExpectData, dg As MSFlexGrid)
    On Error Resume Next
    Dim row As Integer
    Dim maxLong As Integer
    Dim i As Long
    maxLong = IIf(UBound(data) < 80, UBound(data), 80)
    For row = iMinCheckTimes To maxLong
        Dim posDir As New Dictionary
        For i = iColShift + 1 To MaxPos
            Dim NoExitDir As Dictionary
            Dim strDic As Dictionary
            Set strDic = getColumnData(data, i, row, 1, NoExitDir)
            Dim str() As String
            posDir.Add i, strDic
            ReDim str(NoExitDir.Count)
            Dim cnt As Integer
            'Sheet1.Cells(6 + row - iMinCheckTimes, 9 + i) = getNoCode(posDir(i), 1, MaxPos)
            Err.Clear
            dg.TextMatrix(row - iMinCheckTimes + 1, i) = getNoCode(posDir(i), 1, MaxPos)
            If Err.Number <> 0 Then
                 'MsgBox row & ":" & i
                 test = ""
            End If
        Next
        Set posDir = Nothing
    Next
End Sub

Sub SearchCombRecords(combs() As Variant, data() As ExpectData, MinTimes As Integer, p As Integer, ByRef ret() As String)
    On Error Resume Next
    Dim combcnt As Long
    Dim datacnt As Long
    Dim i As Long
    Dim j As Integer
    Dim coli As Integer
    ReDim ret(0)
    combcnt = UBound(combs)
    datacnt = UBound(data)
    Dim posDir As New Dictionary
    For i = iColShift + 1 To MaxPos
        Dim NoExitDir As Dictionary
        posDir.Add i, getColumnData(data, i, MinTimes, 1, NoExitDir)
''        Dim str() As String
''        ReDim str(NoExitDir.Count)
''        Dim cnt As Integer
''        Sheet1.Cells(6 + minTimes - iMinCheckTimes, 9 + i) = getNoCode(posDir(i), 1, MaxPos)
    
    Next
    Dim mcnt As Integer
    For i = 1 To combcnt '遍历p个名次有组合
        '取每名次
        comb = Split(combs(i), ",")
        Dim NoCnt As Integer
        Dim codes() As String
        If UBound(comb) + 1 <> p Then
            GoTo lastRec
        End If
        NoCnt = 0
        For coli = 0 To p - 1
            Dim dir As New Dictionary
            '''sheet4.cells(i, 8 + coli) = "无"
            Set dir = posDir(Int(comb(coli)))
            If dir.Count = MaxPos Then '排名在指定次数里面所有车都出现
                GoTo notColFor
            End If
            NoCnt = NoCnt + 1
            Dim strCode As String
            strCode = getNoCode(dir, 1, MaxPos)
            ReDim Preserve codes(NoCnt)
            codes(NoCnt) = strCode
'''            Sheet4.cells(i, 8 + coli) = strCode
        Next
notColFor:
         Err.Clear
         test = UBound(codes)
         If Err.Number <> 0 Then
            Err.Clear
            GoTo lastRec
         End If
         If UBound(codes) = p Then
            strRet = getSameNoCode(codes, 1, 10)
            If Len(strRet) > 0 Then
'                Sheet4.cells(i, 7) = strRet
'                Sheet4.cells(i, 6) = "存在"
                mcnt = mcnt + 1
                ReDim Preserve ret(mcnt)
                ret(mcnt) = Replace(Replace(combs(i), "10", "0"), ",", "") & "/" & Replace(strRet, ",", "")
                ReDim codes(0)
            End If
        End If
'''        strSame = getSameNoCode(codes, 1, 10)
'''        If Len(Trim(strSame)) > 0 Then
'''            mcnt = mcnt + 1
'''            ReDim Preserve ret(mcnt)
'''            ret(mcnt) = Replace(Replace(combs(i), "10", "0") & "/" & strSame, ",", "")
'''        End If
lastRec:
    Next

    Exit Sub
End Sub

Function getNoCode(dir As Dictionary, iFrom As Integer, iTo As Integer) As String
    On Error Resume Next
    If dir.Count = 0 Then
        getNoCode = ""
        Exit Function
    End If
    Dim i As Integer
    Dim ret() As String
    Dim retcnt As Integer
    For i = iFrom To iTo
        Dim strI As String
        strI = Right("00" & i, 2)
        If Not dir.Exists(strI) Then
            retcnt = retcnt + 1
            ReDim Preserve ret(retcnt)
            ret(retcnt) = i Mod 10
        End If
    Next
    getNoCode = Join(ret, "")
End Function

Function getSameNoCode(codes() As String, iFrom As Integer, iTo As Integer) As String
    On Error Resume Next
    Dim ret() As String
    Dim i As Integer
    Dim j As Integer
    Dim cnt As Integer
    For i = iFrom To iTo
        For j = 1 To UBound(codes)
            Dim sCols As String
            Dim sReplace As String
            sCols = codes(j)
            'sCols = Replace(sCols, "0", "10")
            sReplace = Replace(sCols, IIf(i = 10, "0", CStr(i)), "")
            If sCols = sReplace Then '其中任意一个不符合条件，返回空
                GoTo EndFor
            End If
        Next
        cnt = cnt + 1
        ReDim Preserve ret(cnt)
        ret(cnt) = Right(CStr(i), 1)
EndFor:
    Next
    getSameNoCode = Join(ret, ",")
End Function

Function ReReadHtml() As ExpectData()
    On Error Resume Next
    Dim ret() As ExpectData
    Set httpreq = New XMLHTTP60 ' CreateObject("MSXML2.XMLHTTP.3.0") '
'Dim HttpReq As New MSXML2.XMLHTTP30
    'gobj.LogObj.Log "准备读取最新数据"
    httpreq.Open "get", strurl, False
    'HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    httpreq.send
    'Sheet1.Cells(1, "J") = Now() & "发送数据请求   "
    'Sheet1.Cells(2, "J") = Now() & "接收到数据状态" & HttpReq.Status
    Dim stime As Date
    stime = Now()
    While httpreq.readyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 3 Then
            'getHtmlStr = "OutTime"
            ReDim ret(0)
            ReReadHtml = ret
            Exit Function '判断超出3秒即超时退出过程
        End If
    Wend
    If httpreq.Status <> 200 Then
        ReDim ret(0)
        ReReadHtml = ret
        gobj.LogObj.Log "网页响应错误！"
        Exit Function
    End If
    Dim xmldoc As New DOMDocument60
    Err.Clear
    xmldoc.LoadXML (httpreq.responseXML.XML)
    'gobj.LogObj.Log httpreq.responseXML.XML
    'Sheet3.Cells(1, 1) = Now() & xmldoc.XML
    If Err.Number <> 0 Then
        gobj.LogObj.Log httpreq.responseText
        Exit Function
    End If
    ret = getData(httpreq.responseXML)
    ReReadHtml = ret
End Function


Function ReReadInst() As String
    On Error Resume Next
    Dim ret() As ExpectData
    Dim httpreq As XMLHTTP60
    Set httpreq = New XMLHTTP60 ' CreateObject("MSXML2.XMLHTTP.3.0") '
    Dim strurl As String
    strurl = "http://www.wolfinv.com/pk10/app/requestinsts.asp"
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "If-Modified-Since", "0"
    httpreq.send

    Dim stime As Date
    stime = Now()
    While httpreq.readyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 3 Then
            'getHtmlStr = "OutTime"
            
            ReReadInst = ""
            Exit Function '判断超出3秒即超时退出过程
        End If
    Wend
    If httpreq.Status <> 200 Then
        ReReadInst = ""
        gobj.LogObj.Log "网页响应错误！"
        Exit Function
    End If
    ReReadInst = httpreq.responseText
End Function

Function Login() As String
    On Error Resume Next
    Dim ret() As ExpectData
    Dim httpreq As XMLHTTP60
    Set httpreq = New XMLHTTP60 ' CreateObject("MSXML2.XMLHTTP.3.0") '
    Dim strurl As String
    strurl = "http://www.wolfinv.com/pk10/app/login.asp?User={0}&Password={1}"
    strurl = Replace(Replace(strurl, "{0}", Me.gobj.ClientUserName), "{1}", Me.gobj.ClientPassword)
    httpreq.Open "get", strurl, False
    httpreq.setRequestHeader "If-Modified-Since", "0"
    httpreq.send

    Dim stime As Date
    stime = Now()
    While httpreq.readyState <> 4
        DoEvents
        ntime = Now '获取循环时间
        If DateDiff("s", stime, ntime) > 3 Then
            'getHtmlStr = "OutTime"
            
            Login = ""
            Exit Function '判断超出3秒即超时退出过程
        End If
    Wend
    If httpreq.Status <> 200 Then
        Login = ""
        gobj.LogObj.Log "网页响应错误！"
        Exit Function
    End If
    Login = httpreq.responseText
End Function

Function getData(xmldoc As DOMDocument60)
   Dim datas() As ExpectData
   Set rowNodes = xmldoc.SelectNodes("/xml/row")
   If rowNodes.Length = 0 Then
        'MsgBox "取得数据条数为0"
        getData = datas
        Exit Function
   End If
   Dim firstNode As IXMLDOMNode
   Set firstNode = rowNodes(0)
   Me.NewestExpect = firstNode.Attributes(0).NodeValue
'   If CurrExpect = Sheet1.Cells(1, "B").Text Then '新期号数据已出来
'        Sheet10.Cells(3, 1) = 0
'   End If
'   Sheet1.Cells(1, "B") = CurrExpect + 1
'   Sheet1.Cells(2, "B") = CurrExpect
'   Sheet1.Cells(3, "B") = firstNode.Attributes(2).NodeValue
'   Sheet1.Cells(4, "B") = firstNode.Attributes(1).NodeValue
'   Sheet1.Cells(4, "D") = rowNodes.Length
   Dim i As Integer
   ReDim Preserve datas(rowNodes.Length)
   For i = 1 To rowNodes.Length
      Set datas(i) = New ExpectData
      datas(i).Expect = rowNodes(i - 1).Attributes(0).NodeValue
      datas(i).OpenTime = rowNodes(i - 1).Attributes(2).NodeValue
      datas(i).OpenCode = rowNodes(i - 1).Attributes(1).NodeValue
   Next
   getData = datas
End Function

Function getColumnData(data() As ExpectData, colid As Long, n As Integer, Form As Integer, ByRef NoExitDir As Dictionary) As Dictionary
    On Error Resume Next
    Dim dic As New Dictionary
    Set NoExitDir = New Dictionary
    If n > UBound(data) Then
        Set getColumnData = dic
        Exit Function
    End If
    For i = Form To n
        Dim POS() As String
        Dim strPos As String
        POS = Split(data(i).OpenCode, ",") '排名数组
        strPos = POS(colid - 1)
        If Not dic.Exists(strPos) Then dic.Add strPos, Int(strPos)
    Next
    If Err.Number <> 0 Then
        test = 1
    End If
    For i = 1 To 10
        If Not dic.Exists(Right("00" & i, 2)) Then
            If Not NoExitDir.Exists(Right("00" & i, 2)) Then
                 NoExitDir.Add i, i
            End If
        End If
    Next
    Set getColumnData = dic
End Function

Function getMutliTimesNo()
    
End Function

Sub UpdateComb(cbs() As Variant, ishift As Integer, ByRef ret() As Variant)
    Dim cnt As Long
    cnt = UBound(cbs)
    ReDim ret(cnt)
    Dim i As Long
    For i = 1 To cnt
        subcomb = Split(cbs(i), ",")
        Dim j As Integer
        For j = 0 To UBound(subcomb)
           subcomb(j) = ishift + subcomb(j)
        Next
        ret(i) = Join(subcomb, ",")
    Next

End Sub

Sub WriteComb(combs() As Variant)
    Dim i As Long
    For i = 1 To UBound(combs)
        CombItem = combs(i)
        Dim j As Integer
        'Sheet4.Cells(i, 1) = CombItem '(j)
    Next
End Sub

'获得组合总数
 Private Function Total(ByVal m As Integer, ByVal n As Integer) As Long
 Dim i As Long
 Dim result As Double
 result = 1
 For i = n To 1 Step -1
    result = result * m / i
    m = m - 1
 Next
 Total = result
 'MsgBox Result
End Function

'单个组合
Private Sub CarryComb(CombItem As Variant, pComp() As Long, n As Integer)
 Dim i As Long
 i = n
 While CombItem(i) = pComp(i)
    i = i - 1
 Wend
 CombItem(i) = CombItem(i) + 1
 While i < n
    i = i + 1
    CombItem(i) = CombItem(i - 1) + 1
 Wend
End Sub


'组合结果
Private Sub CombResult(m As Integer, n As Integer, ByRef result() As Variant)
 
 Dim i As Long, j As Long
 Dim Num As Long
 
 If n > (m) Then Exit Sub
 
 ReDim CombItem(1 To n)
 ReDim pComp(1 To n) As Long
 
 Num = m - n
 For i = 1 To n - 1
    CombItem(i) = i
    pComp(i) = Num + i
 Next
 CombItem(n) = n - 1
 pComp(n) = m
 
 Num = Total(m, n)   ': Debug.Print Num
 ReDim result(1 To Num)
 For i = 1 To Num
    CarryComb CombItem, pComp, n
    result(i) = Join(CombItem, ",")
    'DoEvents
 Next
 
End Sub


Sub InitChanceList()
   On Error Resume Next
   Set chanceDic = New Dictionary
    Dim ccs() As ChanceClass
    Dim cc As New ChanceClass
    cc.ReadCurrDateHoldingChances ccs
    Set chanceDic = cc.getHoldChancesDic(ccs)
End Sub
Sub SearchTempColumnInputItems(dgcol As MSFlexGrid, ByRef tmpArr() As String, ByRef allColArr() As String, i As Integer)
    Dim tmpcnt As Integer
    ReDim Preserve tmpArr(0)
    tmpcnt = 0
    cnt = UBound(allColArr)
    If i = 2 Or i = 3 Then
        tmpcnt = 0
    End If
    With dgcol
        Dim j As Integer
        For j = 1 To .Rows
            Dim strCars As String
            strCars = Trim(.TextMatrix(j, i + 1))
            If strCars = "" Then
                maxCar = j + gobj.MutliColMinTimes - 2
                GoTo colend
            End If
            
            Dim k As Integer
            For k = 1 To 8
                If Len(strCars) = k And Len(Trim(.TextMatrix(j + 1, i + 1))) = k - 1 Then
                    '最后一个
                    Dim outTimes As Integer
                    outTimes = j + gobj.MutliColMinTimes - 1
                    If (k > 4 And outTimes >= gobj.MinTimeForChance(k) - 2) Or ((k <= 4) And outTimes >= gobj.MinTimeForChance(k + 2)) Then '对三四列放宽范围为再组合准备数据
                        tmpcnt = tmpcnt + 1
                        ReDim Preserve tmpArr(tmpcnt)
                        tmpArr(tmpcnt) = Replace(Replace(CStr(i + 1), "10", "0") & "/" & strCars & ":" & CStr(outTimes), " ", "")
                        cnt = cnt + 1
                        ReDim Preserve allColArr(cnt)
                        allColArr(cnt) = tmpArr(tmpcnt)
                    End If
                End If
            Next
        Next
colend:
    End With
End Sub

Sub SearchTempAllInputItems(dgAll As MSFlexGrid, ByRef alltmpArr() As String, i As Integer)
    With dgAll
            Dim holdcnt As Integer
            
            Dim alltmpCnt As Integer
            alltmpCnt = 0
            ReDim Preserve alltmpArr(0)
            Dim realCnt As Integer
            realCnt = i
            If realCnt = 2 Then ralcnt = 6
            For j = 1 To .Rows
                'Dim strCars As String
                strCars = Trim(.TextMatrix(j, i - 1))
                If strCars = "" Then
                    holdcnt = j - 1
                    Exit Sub
                End If
                strCars = Trim(strCars)
                strCars = Replace(strCars, "  ", " ")
                strCars = Replace(strCars, "  ", " ")
                strCars = Replace(strCars, "  ", " ")
                strCars = Replace(strCars, "  ", " ")
                Dim currArr() As String
                currArr = Split(strCars, " ")
                Dim nextChances As String
                nextChances = Trim(.TextMatrix(j + 1, i - 1))
                If Len(strCars) > Len(nextChances) Then '如果下个机会小于本次机会
                    'lastchances = Trim(Replace(strCars, nextChances, "")) '截止的机会 ‘此处没考虑到排序变化，必须要逐个消除
                    Dim ni As Integer
                    Dim nextdic As Dictionary
                    Set nextdic = New Dictionary
                    nextChances = Trim(nextChances)
                    nextChances = Replace(nextChances, "  ", " ")
                    nextChances = Replace(nextChances, "  ", " ")
                    nextChances = Replace(nextChances, "  ", " ")
                    nextChances = Replace(nextChances, "  ", " ")
                    Dim nextArr() As String
                    nextArr = Split(nextChances)
                    For ni = 0 To UBound(nextArr)
                        If Not nextdic.Exists(nextArr(ni)) Then
                            nextdic.Add nextArr(ni), nextArr(ni)
                        End If
                    Next
                    Dim lastArr() As String
                    ReDim lastArr(0)
                    Dim li As Integer
                    li = 0
                    ReDim lastArr(0)
                    For ni = 0 To UBound(currArr)
                        If nextdic.Exists(currArr(ni)) Then  '下次不存在，才是真是最后出现的机会
                        Else
                            li = li + 1
                            ReDim Preserve lastArr(li)
                            lastArr(li) = currArr(ni)
                        End If
                    Next
                    lastchances = Trim(Join(lastArr, " "))
                    If Len(lastchances) > 0 Then
                        outTimes = j + gobj.MutliColMinTimes - 1
                        'If outTimes >= gobj.MinTimeForChance(realCnt) - 2 Then
                        If ((i < 3 Or i > 4) And outTimes >= gobj.MinTimeForChance(realCnt) - 2) Or ((i = 3 Or i = 4) And outTimes >= gobj.MinTimeForChance(realCnt + 2) And outTimes > gobj.MutliColMinTimes) Then
                             Dim tchances() As String
                             tchance = Split(lastchances, " ")
                             For k = 0 To UBound(tchance)
                                alltmpCnt = alltmpCnt + 1
                                ReDim Preserve alltmpArr(alltmpCnt)
                                alltmpArr(alltmpCnt) = tchance(k) & ":" & CStr(outTimes)
                             Next
                        End If
                    End If
                End If
            Next
        End With
End Sub

Function getObjectChips(str As String) As Integer
    Dim i As Integer
    Dim objarr() As String
    objarr = Split(str, "+")
    Dim sum As Integer
    sum = 0
    For i = 0 To UBound(objarr)
        Dim sobj As String
        sobj = objarr(i)
        Dim sarr() As String
        sarr = Split(sobj, "/")
        sum = sum + Len(sarr(0)) * Len(sarr(1))
    Next
    getObjectChips = sum
End Function

Sub SearchColumnInputItems(tmpArr() As String, ByRef CurrDic As Dictionary, ByRef retArr() As ChanceClass, MinDiff As Integer)
    Dim cnt As Integer
    cnt = UBound(retArr)
'    Dim CurrDic As Dictionary
'    Set CurrDic = New Dictionary
    For j = 1 To UBound(tmpArr)
        Dim tmpstr As String
        tmpstr = tmpArr(j)
        Dim strChance As String
        Dim iTimes As Integer
        Dim MinInTimes As Integer
        Dim strLen As Integer
        strChance = ""
        strChance = Split(tmpstr, ":")(0)
        'strLen = Len(strChance) - 2 '此情况只针对单一车，名词
        '必须要按组合来
        strLen = getObjectChips(strChance)
        If strLen <= 0 Then GoTo ColNextEnd
'''''                    MinInTimes = gobj.MinTimeForChance(strLen) '基础次数
        MinInTimes = gobj.MinTimeForChance(strLen) + MinDiff '基础次数,如果当日所有记录数小于检查下限，增加2
        Dim tMinInTimes As Integer
        tMinInTimes = MinInTimes
        
        Dim arr() As String
        arr = Split(strChance, "/")
        iTimes = CInt(Split(tmpstr, ":")(1))
        Dim chips As Integer
        chips = Me.getObjectChips(strChance)
        Dim singleAddTimes As Integer
        singleAddTimes = IIf(chips >= 8, -3, 8 - chips)
        If chips <= 4 Then '对于注数为3或4的组合
            If Not CurrDic.Exists(strChance) Then
                CurrDic.Add strChance, CInt(iTimes)
            Else
                If CInt(iTimes) > CurrDic(strChance) Then
                    CurrDic(strChance) = CInt(iTimes) '如果同一机会持续次数高于存在机会，更新机会的次数
                End If
            End If
        End If
        If chanceDic.Exists(strChance) Then '如果往期已存在，不做判断，直接加入机会清单
            cnt = cnt + 1
            ReDim Preserve retArr(cnt)
            Set retArr(cnt) = New ChanceClass
            retArr(cnt).ChanceCode = strChance
            retArr(cnt).ExpectCode = currExpect
            retArr(cnt).ChanceType = chanceDic(strChance).ChanceType
            retArr(cnt).BaseCost = chanceDic(strChance).BaseCost
        Else
            If maxCar > ISingleMin And maxCar < ISingleMax Then '如果其中最大持续未出现车次超过35，进入条件增加2
                MinInTimes = MinInTimes + 2
            End If
            needAddCnt = 0
            needAddCnt = Me.LikeAllBorS(CStr(Split(strChance, "/")(1))) ' Right(strChance, Len(strChance) - 2))
            MinInTimes = MinInTimes + needAddCnt
            needAddCnt = Me.LikeAllJorO(CStr(Split(strChance, "/")(1))) 'Right(strChance, Len(strChance) - 2))
            MinInTimes = MinInTimes + needAddCnt
            needAddCnt = 0
            If MinInTimes < tMinInTimes + singleAddTimes Then
                MinInTimes = tMinInTimes + singleAddTimes
            End If
            If iTimes >= MinInTimes Then
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChance
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = iTimes
            End If
        End If
ColNextEnd:
    Next

End Sub

Sub SearchAllInputItems(alltmpArr() As String, ByRef retArr() As ChanceClass, MinInTimes As Integer, LongHoldingDic As Dictionary)
       Dim cnt As Integer
       cnt = UBound(retArr)
       Dim CurrDic As Dictionary
       Dim tMinInTimes As Integer
       tMinInTimes = MinInTimes
       Set CurrDic = New Dictionary
        For j = 1 To UBound(alltmpArr)
            Dim txtchance As String
            txtchance = Trim(Split(alltmpArr(j), ":")(0))
            Dim arr() As String
            arr = Split(txtchance, "/")
            iTimes = Split(alltmpArr(j), ":")(1)
            Dim chips As Integer
            chips = Me.getObjectChips(txtchance)
            If chips = 3 Or chips = 4 Then '对于长度为3或4的组合
                If Not CurrDic.Exists(txtchance) Then
                    CurrDic.Add txtchance, CInt(iTimes)
                Else
                    If CInt(iTimes) > CurrDic(txtchance) Then
                        CurrDic(txtchance) = CInt(iTimes) '如果同一机会持续次数高于存在机会，更新机会的次数
                    End If
                End If
            End If
            For k = 0 To Len(arr(0)) - 1
                tt = Mid(arr(0), k + 1, 1) & "/" & arr(1)
                If LongHoldingDic.Exists(tt) Then
                    MinInTimes = MinInTimes + IIf(Len(arr(0)) <= 4, 2, 1)
                End If
            Next
            Dim singleAddTimes As Integer
            singleAddTimes = IIf(isPChance(txtchance), IIf(chips >= 8, -3, 8 - chips), 0)
            If MinInTimes < tMinInTimes + singleAddTimes Then
                MinInTimes = tMinInTimes + singleAddTimes
            End If
            If chanceDic.Exists(txtchance) Then '如果往期已存在，不做判断，直接加入机会清单
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = txtchance
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).ChanceType = chanceDic(txtchance).ChanceType
                retArr(cnt).BaseCost = chanceDic(txtchance).BaseCost
            Else
                
                If iTimes >= MinInTimes Then
                    cnt = cnt + 1
                    ReDim Preserve retArr(cnt)
                    Set retArr(cnt) = New ChanceClass
                    retArr(cnt).ChanceCode = txtchance
                    retArr(cnt).ExpectCode = currExpect
                    retArr(cnt).InputTimes = iTimes
                End If
            End If
        Next
            '获得组合的组合
        Me.SearchMixableItems CurrDic, retArr, False
End Sub

Function isPChance(strCode As String)
    If Len(Trim(strCode)) > 3 Then
        Dim arr() As String
        arr = Split(strCode, "/")
        If (Len(Trim(arr(0))) > 1 And Len(Trim(arr(1))) > 1) Then
            isPChance = False
            Exit Function
        End If
    End If
    isPChance = True
End Function

Function SearchMixableItems(dic As Dictionary, ByRef retArr() As ChanceClass, isColumn As Boolean)
    Dim cnt As Integer
    cnt = UBound(retArr)
    Dim mixedDic As New Dictionary
    Dim tmpdic As New Dictionary
    Set tmpdic = dic
    Dim arr() As String
    ReDim arr(0)
    Dim acnt As Integer
    For Each k In dic.Keys
        If Not IsEmpty(k) Then
            acnt = acnt + 1
            ReDim Preserve arr(acnt): arr(acnt) = k
        End If
    Next
    For acnt = 1 To UBound(arr)
        k1 = arr(acnt)
        Dim ck1 As Integer
        Dim cTime1 As Integer
        Dim chips As Integer
        Dim strK1 As String
        strK1 = k1
        chips = getObjectChips(strK1)
        ck1 = CInt(Split(k1, "/")(IIf(isColumn, 0, 1)))
        cTime1 = dic(k1)
        Dim j As Integer
        For j = 1 To UBound(arr)
            k2 = arr(j)
            Dim ck2 As Integer
            Dim cTime2 As Integer
            Dim strK2 As String
            strK2 = k2
            cTime2 = dic(k2)
            ck2 = CInt(Split(k2, "/")(IIf(isColumn, 0, 1)))
            If Len(k1) <> Len(k2) Then
                GoTo repeatContinue
            End If
            If k1 = k2 Then
                GoTo repeatContinue
            End If
            If ck1 = ck2 Then
                GoTo repeatContinue
            End If
            Dim RepeatCnt As Integer
            Dim strMixKeys As String
            Dim strChanceCode As String
            RepeatCnt = getRepeatCnt(strK1, strK2, isColumn)
            strMixKeys = IIf(ck2 > ck1, IIf(isColumn, "", "|") & CStr(ck1) & IIf(isColumn, "", "|") & CStr(ck2), IIf(isColumn, "", "|") & CStr(ck2) & IIf(isColumn, "", "|") & CStr(ck1)) & "-" & CStr(Len(k1) - 2)
            strChanceCode = IIf(ck2 > ck1, k1 & "+" & k2, k2 & "+" & k1)
            If mixedDic.Exists(strMixKeys) Then GoTo repeatContinue
            Dim MinTimesF2, MinTimesF1, MinTimes, MinTimesF15, MinTimesF05 As Integer
            MinTimesF2 = gobj.MinTimeForChance(chips + 2)
            MinTimesF1 = gobj.MinTimeForChance(chips + 1)
            MinTimesF15 = (MinTimesF1 + MinTimesF2) / 2
            MinTimes = gobj.MinTimeForChance(chips)
            MinTimesF05 = (MinTimesF1 + MinTimes) / 2
            If RepeatCnt = 4 And cTime1 >= MinTimesF2 Then '如果大于3个相同的，不用任何判断
                mixedDic.Add strMixKeys, strChanceCode
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChanceCode
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = IIf(cTime1 < cTime2, cTime1, cTime2)
                GoTo repeatContinue
            End If
            If RepeatCnt >= 3 And cTime1 >= MinTimesF2 And cTime2 >= MinTimesF2 And cTime1 + cTime2 >= 2 * MinTimesF15 Then 'cTime1 >= MinTimesF15 And cTime2 >= MinTimesF15 And cTime1 + cTime2 >= 2 * MinTimesF1 Then '如果等于3个相同的，两数均要大于前1，2未中值
                mixedDic.Add strMixKeys, strChanceCode
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChanceCode
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = IIf(cTime1 < cTime2, cTime1, cTime2)
                GoTo repeatContinue
            End If
            If RepeatCnt >= 2 And cTime1 >= MinTimesF15 And cTime2 >= MinTimesF15 And cTime1 + cTime2 >= 2 * MinTimesF1 Then 'cTime1 >= MinTimesF1 And cTime2 >= MinTimesF1 And cTime1 + cTime2 >= 2 * MinTimesF05 Then '两位需要大于长一位需要的最小长度
                mixedDic.Add strMixKeys, strChanceCode
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChanceCode
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = IIf(cTime1 < cTime2, cTime1, cTime2)
                GoTo repeatContinue
            End If
            If RepeatCnt >= 1 And cTime1 >= MinTimesF1 And cTime2 >= MinTimesF1 And cTime1 + cTime2 >= 2 * MinTimesF05 Then 'cTime1 >= MinTimesF1 And cTime2 >= MinTimesF1 And cTime1 + cTime2 >= 2 * MinTimes Then  '两者次数要超过长一位组合需要长度，并且两者相加长度大于本位及长一位中值
                mixedDic.Add strMixKeys, strChanceCode
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChanceCode
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = IIf(cTime1 < cTime2, cTime1, cTime2)
                GoTo repeatContinue
            End If
            If RepeatCnt = 0 And cTime1 >= MinTimesF05 And cTime2 >= MinTimesF05 Then  '两者次数要超过本位组合需要长度，并且两者相加长度均值大于本位
                mixedDic.Add strMixKeys, strChanceCode
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = strChanceCode
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = IIf(cTime1 < cTime2, cTime1, cTime2)
                GoTo repeatContinue
            End If
repeatContinue:
        Next

    Next
End Function

Function getRepeatCnt(k1 As String, k2 As String, isColumn As Boolean) As Integer
    Dim str1, str2, strAll As String
    Dim dic As New Dictionary
    str1 = Split(k1, "/")(IIf(isColumn, 1, 0))
    str2 = Split(k2, "/")(IIf(isColumn, 1, 0))
    strAll = str1 & str2
    Dim i As Integer
    For i = 1 To Len(strAll)
        If Not dic.Exists(Mid(strAll, i, 1)) Then
            dic.Add Mid(strAll, i, 1), 1
        End If
    Next
    getRepeatCnt = Len(strAll) - dic.Count
    
    
    
''    If isColumn Then
''        'str1 = right(str1,len(str1)-2)
''        str1 = Split(k1, "/")(1)
''        str2 = Split(k2, "/")(1)
''    Else
''    End If
End Function

Function SearchRepeatItems(frm As Form, ByRef retArr() As ChanceClass)
    Dim cnt As Integer
    cnt = UBound(retArr)
    Dim txtRepeatChancesCtrl As TextBox
    Dim txtRepeatCntCtrl As TextBox
    Set txtRepeatChancesCtrl = CallByName(frm, "txt_RepeateChances", VbGet)
    Set txtRepeatCntCtrl = CallByName(frm, "txt_RepeatCnt", VbGet)
    '此处存在风险，因历史数据取回来不及时，后面取回来的历史数据会影响排序
    Dim strChances, strCnts As String
    strChances = Replace(Trim(txtRepeatChancesCtrl.Text), "  ", " ")
    strChances = Replace(strChances, "  ", " ")
    strChances = Replace(strChances, "  ", " ")
    strChances = Replace(strChances, "  ", " ")
    strChances = Replace(strChances, "  ", " ")
    strChances = Replace(strChances, "  ", " ")
    strCnts = Trim(txtRepeatCntCtrl.Text)
    Dim Scnt, Ccnt As Integer
    Scnt = Split(strCnts, "/")(0)
    Ccnt = Split(strCnts, "/")(1)
    Dim checkScnt, checkCcnt As Integer
    Dim checkCntArr() As String
    checkCntArr = Split(gobj.RepeatCheckCnt, "/")
    checkScnt = checkCntArr(0)
    checkCcnt = checkCntArr(1)
    If Len(Trim(strChances)) > 0 And CInt(Scnt) > CInt(checkScnt) And CInt(Ccnt) > CInt(checkCcnt) Then
        Dim strRArr() As String
        Dim strPlus As String
        strRArr = Split(strChances, " ")
        strPlus = Replace(strChances, " ", "+")
        cnt = cnt + 1
        ReDim Preserve retArr(cnt)
        Set retArr(cnt) = New ChanceClass
        retArr(cnt).ChanceCode = strPlus
        retArr(cnt).ChanceType = 1
        retArr(cnt).ExpectCode = currExpect & ":" & Trim(strCnts)
        retArr(cnt).InputTimes = CInt(Scnt)
        retArr(cnt).ChipCount = getObjectChips(strPlus)
    End If
End Function

'获得所有可能可以下注的机会
Public Sub SearchInputItems(frm As Form, dgAll As MSFlexGrid, dgcol As MSFlexGrid, chanceDic As Dictionary, ByRef ret() As ChanceClass)
    
    On Error Resume Next
    Dim retArr() As ChanceClass
    If ExpectChancesDic.Exists(currExpect) Then
        ret = ExpectChancesDic(currExpect)
        Exit Sub
    End If
    ReDim Preserve retArr(0)
    Dim cnt As Integer
    Dim i As Integer
    Dim singleChances(10) As String
    Dim LongHoldingDic As New Dictionary
    Dim mutliLongArr() As String
    Dim mi As Integer
    For i = gobj.StartCols To 9
        Dim txtctrl As TextBox
        Dim arg(1) As Variant
        arg(0) = i
        Err.Clear
        Set txtctrl = CallByName(frm, "txt_SingleCol", VbGet, i)
        If Err.Number <> 0 Then
            'MsgBox Err.Description
            test = ""
        End If
        singleChances(i) = txtctrl.Text
        If (Len(Trim(txtctrl.Text)) > 0) Then
            arr = Split(singleChances(i), ":")
            If arr(1) >= ISingleMin And arr(1) < ISingleMax Then
                LongHoldingDic.Add arr(0), arr(1)
            End If
            If arr(1) > gobj.MinTimeForChance(1) Then
                cnt = cnt + 1
                ReDim Preserve retArr(cnt)
                Set retArr(cnt) = New ChanceClass
                retArr(cnt).ChanceCode = arr(0)
                retArr(cnt).ExpectCode = currExpect
                retArr(cnt).InputTimes = arr(1)
            End If
            If arr(1) > gobj.SingleColMinTimes Then '黄色单车
                mi = mi + 1
                ReDim Preserve mutliLongArr(mi)
                mutliLongArr(mi) = arr(0)
            End If
        End If
    Next
    If mi > 2 Then '大于黄色的所有集合
        cnt = cnt + 1
        ReDim Preserve retArr(cnt)
        Set retArr(cnt) = New ChanceClass
        retArr(cnt).ChanceCode = Replace(Trim(Join(mutliLongArr, " ")), " ", "+")
        retArr(cnt).ExpectCode = currExpect
        retArr(cnt).InputTimes = gobj.SingleColMinTimes + 1
        retArr(cnt).ChipCount = mi
    End If
    Dim MinDiff As Integer
    MinDiff = IIf(Me.CurrDataCount < ISingleMin, 2, 0)
    'I = gobj.StartCols
    Dim allColArr() As String
    Dim allColDic As New Dictionary
    ReDim allColArr(0)
    For i = gobj.StartCols To 10 - 1
        Dim maxCar As Integer
        Dim tmpArr() As String
        SearchTempColumnInputItems dgcol, tmpArr, allColArr, i
        If UBound(tmpArr) > 0 Then SearchColumnInputItems tmpArr, allColDic, retArr, MinDiff '获得车道组合
    Next
    '获得列组合的组合
    Me.SearchMixableItems allColDic, retArr, True
    For i = 2 To 8
        Dim alltmpArr() As String
        Dim realCnt As Integer
        realCnt = i
        If realCnt = 2 Then ralcnt = 6
        Me.SearchTempAllInputItems dgAll, alltmpArr, i
        Dim MinInTimes As Integer
''''''        MinInTimes = gobj.MinTimeForChance(I) '基础次数
        '单项要增加入场次数

        MinInTimes = gobj.MinTimeForChance(realCnt) + MinDiff  '基础次数,如果当日所有记录数小于检查下限，增加2
        SearchAllInputItems alltmpArr, retArr, MinInTimes, LongHoldingDic
    Next
    
    
    
    SearchRepeatItems frm, retArr
    
    
    
    
    Dim dicCc As Dictionary
    Set dicCc = New Dictionary
    ReDim ret(0)
    Dim retcnt As Integer
    retcnt = 0
    For i = 1 To UBound(retArr)
        Dim cc As ChanceClass
        Set cc = retArr(i)
        If Not dicCc.Exists(Trim(cc.ChanceCode)) Then
            dicCc.Add Trim(cc.ChanceCode), cc.ChanceCode
            retcnt = retcnt + 1
            ReDim Preserve ret(retcnt)
            Set ret(retcnt) = New ChanceClass
            ret(retcnt).ChanceCode = Trim(retArr(i).ChanceCode)
            ret(retcnt).ExpectCode = retArr(i).ExpectCode
            ret(retcnt).InputTimes = retArr(i).InputTimes
            ret(retcnt).ChanceType = retArr(i).ChanceType
        End If
    Next
    For Each key In chanceDic.Keys '上期遗留的未平仓持续组合加入下期组合，防止条件改变，有漏掉组合
        Dim strKey As String
        strKey = key
        Dim mcnt As Integer
        If dicCc.Exists(strKey) = False And chanceDic(key).ChanceType <> 1 And Me.JudgeChanceIsMatch(chanceDic(key), mcnt) = False Then
            retcnt = retcnt + 1
            ReDim Preserve ret(retcnt)
            Set ret(retcnt) = New ChanceClass
            ret(retcnt).ChanceCode = Trim(chanceDic(key).ChanceCode)
            ret(retcnt).ExpectCode = chanceDic(key).ExpectCode
            ret(retcnt).InputTimes = chanceDic(key).InputTimes
            ret(retcnt).ChanceType = chanceDic(key).ChanceType
            ret(retcnt).BaseCost = chanceDic(key).BaseCost
        End If
    Next
    If Not Me.ExpectChancesDic.Exists(currExpect) Then
        Me.ExpectChancesDic.Add currExpect, ret
    End If
    Exit Sub
    
''''''    '1
''''''    'road single chance
''''''    Dim OneCarChances As String
''''''    OneCarChances = Trim(strAllSingleChances)
''''''    If Len(OneCarChances) > 0 Then
''''''        Dim strItem() As String
''''''        strItem = Split(OneCarChances, " ")
''''''        For i = 0 To UBound(strItem) '遍历所有单个机会
''''''            item = strItem(i)
''''''            Line = Split(item, "/")(0) '名次
''''''            strline = Line
''''''            If Line = 0 Then strline = 10
''''''            car = Split(item, "/")(1) '车
''''''            chance = Trim(Line & "/" & car)
''''''            cnt = cnt + 1
''''''            ReDim Preserve retArr(cnt)
''''''            Set retArr(cnt) = New ChanceClass
''''''            retArr(cnt).ChanceCode = chance
''''''            retArr(cnt).ExpectCode = CurrExpect
''''''            Dim txtbox As TextBox
''''''            Set txtbox = frm.txt_SingleCol(CInt(strline) - 1)
''''''            currFullInfo = txtbox.Text ' Sheet1.cells(4, 9 + CInt(strline))
''''''            retArr(cnt).CurrTimes = Split(currFullInfo, ":")(1)
''''''        Next
''''''    End If
   
    
''''    'mutli road chances
''''    For i = 3 To 6
''''        Dim txtChances As String
''''        txtChances = Sheet1.cells(2, i + 1)
''''        If Len(Trim(txtChances)) > 0 Then
''''            arr = Split(txtChances, ":") '1234/5 2345/6:25形式
''''            Times = arr(1)
''''            Items = arr(0)
''''            ccs = Split(Trim(Items))
''''            For j = 0 To UBound(ccs)
''''                cnt = cnt + 1
''''                ReDim Preserve retArr(cnt)
''''                Set retArr(cnt) = New ChanceClass
''''                retArr(cnt).ChanceCode = ccs(j)
''''                retArr(cnt).CurrTimes = Times
''''                retArr(cnt).ChipCount = i * Len(Split(ccs(j), "/")(1))
''''                retArr(cnt).ExpectCode = CurrExpect
''''            Next
''''
''''        End If
''''    Next
End Sub

Function LikeAllBorS(strChance As String) As Integer '是否大小
    If Len(strChance) = 3 Then
        If strChance = "123" Or strChance = "124" Or strChance = "125" Or strChance = "134" Or strChance = "135" Or strChance = "145" Or strChance = "234" Or strChance = "235" Or strChance = "245" Or strChance = "345" Then
            LikeAllBorS = 2
            Exit Function
        End If
        If strChance = "678" Or strChance = "679" Or strChance = "670" Or strChance = "689" Or strChance = "680" Or strChance = "690" Or strChance = "789" Or strChance = "780" Or strChance = "790" Or strChance = "890" Then
            LikeAllBorS = 2
            Exit Function
        End If
        Exit Function
    End If
    Dim bcnt As Integer
    Dim Scnt As Integer
    For i = 0 To Len(strChance) - 1
        Dim iChar As Integer
        iChar = Mid(strChance, i + 1, 1)
        If iChar >= 1 And iChar < 6 Then
            bcnt = bcnt + 1
        Else
            Scnt = Scnt + 1
        End If
    Next
    If bcnt <= 1 Or Scnt <= 1 Then '如果全不是属于单双/大小，那要增加更多 ，4码增加3，5码增2，6码增1
        LikeAllBorS = IIf(bcnt = 0 Or Scnt = 0, 7 - Len(strChance), 1)
    End If
End Function

Function LikeAllJorO(strChance As String) As Integer '是否单双
    If Len(strChance) = 3 Then
        If strChance = "135" Or strChance = "137" Or strChance = "139" Or strChance = "157" Or strChance = "159" Or strChance = "179" Or strChance = "357" Or strChance = "359" Or strChance = "379" Or strChance = "579" Then
            LikeAllJorO = 2
            Exit Function
        End If
        If strChance = "246" Or strChance = "248" Or strChance = "240" Or strChance = "268" Or strChance = "260" Or strChance = "280" Or strChance = "468" Or strChance = "460" Or strChance = "480" Or strChance = "680" Then
            LikeAllJorO = 2
            Exit Function
        End If
        Exit Function
    End If
    Dim bcnt As Integer
    Dim Scnt As Integer
    For i = 0 To Len(strChance) - 1
        Dim iChar As Integer
        iChar = Mid(strChance, i + 1, 1)
        If iChar Mod 2 = 1 Then
            bcnt = bcnt + 1
        Else
            Scnt = Scnt + 1
        End If
    Next
    If bcnt <= 1 Or Scnt <= 1 Then
        LikeAllJorO = IIf(bcnt = 0 Or Scnt = 0, 7 - Len(strChance), 1)
    End If
End Function



'获取当前的下注信息
Public Function getCurrExpectChips(frm As Form, allgrid As MSFlexGrid, colgrid As MSFlexGrid) As String
    On Error Resume Next
    Dim ccs() As ChanceClass
    ReDim ccs(0)
    LoadAssetInfo
    Set DicFClose = New Dictionary
    InitChanceList '重新获取新的持有机会表,排除已关闭机会
    SearchInputItems frm, allgrid, colgrid, chanceDic, ccs
    Me.CloseTheEndChances ccs, False '先关闭已结束机会
    InitChanceList
    ReDim ccs(0)
   
    SearchInputItems frm, allgrid, colgrid, chanceDic, ccs
    Dim strResult As String
    strResult = getChips(ccs, True) '第一次下注,强平掉超出止损部分，继续下注部分以及新下注部分
    getCurrExpectChips = strResult
    Dim frm1 As Form1
    Set frm1 = frm
    Set frm1.Gobalobj = Me.gobj
    frm1.Gobalobj.lastExpect = gobj.lastExpect
    Dim retamt As Currency
    frm1.SendMsg Me.gobj.lastExpect, getCurrExpectChips
    'frm1.StatusBar1.Panels(2).Text = retamt
    If Len(Trim(getCurrExpectChips)) > 0 Then gobj.LogObj.OutRecord gobj.lastExpect, Me.ToSerial(getCurrExpectChips, gobj.MinChips)
    If CurrTimeHaveFClose Then '如果有强平的机会,再次载入,重新买入强平机会,不受时间限制
        InitChanceList '重新获取新的持有机会表，排除强平机会
        strResult = getChips(ccs, False)  '再次下注强平机会，因其次数远远超出随机数，故纯粹是继续买入，从头开始
        getCurrExpectChips = strResult
        frm1.SendMsg gobj.lastExpect, strResult
        'frm1.StatusBar1.Panels(2).Text = retamt
        If Len(Trim(getCurrExpectChips)) > 0 Then gobj.LogObj.OutRecord gobj.lastExpect, Me.ToSerial(getCurrExpectChips, gobj.MinChips)
        Set DicFClose = Nothing
    End If
    Me.UpdateAsset
End Function

Function CloseTheEndChances(ccs() As ChanceClass, fclose As Boolean)
    On Error Resume Next
     '先要强制检查是否是最新数据
     If UBound(NewestData) > 0 Then
        If Now() > DateAdd("n", 5, NewestData(1).OpenTime) Then '如果当前的时间比收到的最新数据的开盘时间大于5分钟，暂时不要平，因为最新数据还没收到
            Me.gobj.LogObj.Log "强制平仓发现当前时间超过最新开出数据5分钟，暂时不平仓！"
            Exit Function
        End If
     End If
     For Each key In chanceDic.Keys
        Dim fmatch As Boolean
        Dim i As Integer
        fmatch = False
        Dim cc As ChanceClass
        Set cc = chanceDic(key)
        Dim matchCnt As Integer
        For i = 1 To UBound(ccs)
            If ccs(i).ChanceCode = key Then '如果选出来的组合里面有持仓组合，不平仓
                'fmatch = True
                GoTo breakExec '不处理
            End If
        Next
        matchCnt = 0
        If JudgeChanceIsMatch(cc, matchCnt) = True Then '是否中奖
                fmatch = True
                CloseAChance cc, matchCnt, fclose
        End If
        If matchCnt = -1 Then GoTo breakExec
        If cc.ChanceType = 1 Then '如果是一次性机会，无论是否中奖了，都关闭机会
            CloseAChance cc, matchCnt, fclose
        End If
breakExec:
     Next
     Me.UpdateAsset
End Function
'支持 123/25+1233/25 形式
'cc 组合
'返回命中码个数
Function JudgeChanceIsMatch(cc As ChanceClass, Optional ByRef MatchTimes = -1) As Boolean
    On Error Resume Next
    JudgeChanceIsMatch = False
    Dim OpenCode As String
    Dim lngExpNo As Long
    lngExpNo = CLng(gobj.NewestExpectData.Expect)
    OpenCode = gobj.NewestExpectData.OpenCode
    Dim lastExp As Boolean
    If CLng(cc.ExpectCode) + cc.HoldTimeCnt <> lngExpNo Then
        '去寻找到匹配的期数
        If CLng(cc.ExpectCode) + cc.HoldTimeCnt > lngExpNo Then '还未开出，不处理
            MatchTimes = -1
            Exit Function
        Else
            lastExp = True
            OpenCode = ""
            Dim ecnt As Integer
            ecnt = lngExpNo - CLng(cc.ExpectCode) '防止后面数据缺失，从下注期开始找
            Dim CurrData() As ExpectData
            CurrData = gobj.CurrExpectData
            Dim ec As Integer
            For ec = 0 To ecnt
                Dim testExp As ExpectData
                Set testExp = CurrData(ecnt - ec + 1)
                If testExp Is Nothing Then
                    JudgeChanceIsMatch = True '过期了，必须要强制平掉
                    MatchTimes = 0
                    gobj.LogObj.Log cc.ChanceCode & "已过期，强制中奖0元！"
                    Exit Function '找到最近一个元素了
                Else
                    If CLng(testExp.Expect) = CLng(cc.ExpectCode) + cc.HoldTimeCnt Then  '匹配到对应的开奖期数了
                        OpenCode = testExp.OpenCode
                        GoTo breakFor
                    End If
                End If
            Next
breakFor:
        End If
    End If
    If OpenCode = "" Then
        JudgeChanceIsMatch = True '过期了，必须要强制平掉
        MatchTimes = -1
        gobj.LogObj.Log cc.ChanceCode & "无法找到中奖期号！"
        Exit Function
    End If
    Dim strSerNos, strCars As String
    Dim SCArr() As String
    Dim i As Integer
    Dim plusArr() As String
    plusArr = Split(Trim(cc.ChanceCode), "+")
    Dim matchCnt As Integer
    matchCnt = 0
    For i = 0 To UBound(plusArr)
        SCArr = Split(Trim(plusArr(i)), "/")
        strCars = SCArr(0)
        strSerNos = SCArr(1)
        Dim j As Integer
        For j = 1 To Len(strCars)
            Dim strCar, strSerNo As String
            strCar = Mid(strCars, j, 1)
            strCar = IIf(strCar = "0", "10", strCar)
            strSerNo = Split(OpenCode, ",")(CInt(strCar) - 1) '当期车道所对应名次
            strSerNo = Right(strSerNo, 1) '10取0
            If (InStr(1, strSerNos, strSerNo) > 0) Then '如果用对应的名词在组合的名词组，说明命中
                matchCnt = matchCnt + 1
            End If
        Next

    Next
    If MatchTimes = 0 Then
        MatchTimes = matchCnt
    End If
    If lastExp Then
        JudgeChanceIsMatch = True '过去期数，无论是否中奖都要关闭
        Exit Function
    End If
    If matchCnt > 0 Then JudgeChanceIsMatch = True
End Function

Function getRevChance(strChance As String) As String
    Dim arr() As String
    Dim isCol As Boolean
    arr = Split(strChance, "/")
    isCol = (Len(arr(0)) = 1)
    Dim strModel As String
    strModel = "1234567890"
    Dim strCars As String
    strCars = Trim(IIf(isCol, arr(1), arr(0)))
    Dim i As Integer
    Dim strRet As String
    strRet = strModel
    For i = 1 To Len(strCars)
        strRet = Replace(strRet, Mid(strCars, i, 1), "")
    Next
    getRevChance = IIf(isCol, arr(0) & "/" & strRet, strRet & "/" & arr(1))
End Function

Sub CloseAChance(cc As ChanceClass, cnt As Integer, fclose As Boolean)
    On Error Resume Next
    cc.IsEnd = 1
    cc.CurrTimes = cc.InputTimes + cc.HoldTimeCnt - 1
    If (cc.ChanceType = 1) Then
        cc.CurrTimes = cc.InputTimes - FirstRepeatCnt + 1
    End If
    
    If fclose = True Then
        cc.Profit = cc.Gained - cc.Cost
        Assetobj.Gained = Assetobj.Gained - cc.Cost
        cc.Save 'cc.updateExcel '需要更新
        Exit Sub
    End If
    If cc.ChanceType = 1 And cnt > 0 Then '结束重复号码跟踪,第一个重复的号码数归零
        FirstRepeatCnt = 0
    End If
    If (cc.ChanceType <> 2) Then
        cc.Gained = gobj.Odds * cc.UnitCost * cnt
    Else
        cc.Gained = cc.Cost
    End If
     Assetobj.Gained = Assetobj.Gained + cc.Gained - cc.Cost
     'Assetobj.CurrCash = Assetobj.CurrCash + cc.Gained
     Assetobj.TotalAsset = Assetobj.TotalAsset + cc.Gained
     If cc.ChipCount = 0 Then Me.getObjectChips (cc.ChanceCode)
     If cc.Cost = 0 Then cc.Cost = cc.UnitCost * cc.ChipCount
     cc.Profit = cc.Gained - cc.Cost
     cc.Save ' cc.updateExcel 需要更新
End Sub

 Function getChips(ccs() As ChanceClass, fclose As Boolean) As String
     On Error Resume Next
    If fclose = True And Not Assetobj.AbsLossCheck() Then '如果所有合计最大亏损已经超过了绝对亏损比例，不再下单,并且关闭所有的在持下注
        Dim testccs(0) As ChanceClass
        Me.CloseTheEndChances ccs, True
        getChips = ""
        Exit Function
    End If
    Dim ret() As String
    Dim cnt As Integer
    For i = 1 To UBound(ccs)
        Dim cc As ChanceClass
        Set cc = ccs(i)
        'If cc.ChipCount > 5 Then GoTo ContinueFor '6,8,9,10,12 guolv
        Dim btest As Boolean
        Dim btest1 As Boolean
        btest = chanceDic.Exists(cc.ChanceCode)
        Dim existChance As ChanceClass
        If btest Then
            Set existChance = chanceDic(cc.ChanceCode)
        End If
        btest1 = existChance Is Nothing
        If btest And Not btest1 Then     '如果已经下注过了
            'If cc.ExpectCode + cc.HoldTimeCnt > CurrExpect + 1 Then
            If (cc.ChanceType <> 1) Then
                If chanceDic(cc.ChanceCode).ExpectCode + chanceDic(cc.ChanceCode).HoldTimeCnt > currExpect Then
                '防止重复下注
                    GoTo ContinueFor
                End If
            Else
                GoTo ContinueFor
            End If
            
            If fclose And cc.ChanceType = 0 Then '如果需要强平
                ctimes = cc.CurrTimes
                Set cc = chanceDic(cc.ChanceCode) '获得持有信息
                Dim testcc As New ChanceClass
                testcc.ExpectCode = cc.ExpectCode
                testcc.ChanceCode = cc.ChanceCode
                testcc.ChipCount = IIf(cc.ChipCount = 0, getObjectChips(cc.ChanceCode), cc.ChipCount)
                testcc.HoldTimeCnt = cc.HoldTimeCnt + 1
                testcc.UnitCost = getChanceChips(testcc, currExpect)
                testcc.Cost = cc.Cost + testcc.UnitCost * testcc.ChipCount '用未来成本作为风控成本
                If Assetobj.AllowContinueHold(testcc) Then  '单个机会未超出允许范围,可以继续下单
                    cc.HoldTimeCnt = cc.HoldTimeCnt + 1
                    If cc.ChanceType = 0 Then
                        
                        If cc.ChipCount <= 8 Then
                            cc.UnitCost = getChanceChips(cc, currExpect)
                        Else
                            cc.UnitCost = 14 ^ cc.HoldTimeCnt '八码以上以14的平方递增
                        End If
                        cc.Cost = cc.Cost + cc.UnitCost * cc.ChipCount
                    End If
                    cc.Gained = 0
                    cc.CurrTimes = ctimes
                    cc.Save 'cc.updateExcel 需要更新
                    cnt = cnt + 1
                    Assetobj.CurrCash = Assetobj.CurrCash + cc.UnitCost * cc.ChipCount
                    Assetobj.TotalAsset = Assetobj.TotalAsset - cc.UnitCost * cc.ChipCount
                    ReDim Preserve ret(cnt)
                    ret(cnt) = Replace(cc.ChanceCode, "+", "/" & CStr(cc.UnitCost * gobj.SerTotal(cc.ChipCount)) & "+") & "/" & CStr(cc.UnitCost * gobj.SerTotal(cc.ChipCount))
                Else
                    CloseAChance chanceDic(cc.ChanceCode), 0, True '强平命中次数为0
                    CurrTimeHaveFClose = True '执行完后再次载入机会下注信息，重新执行下注
                    If Not DicFClose.Exists(cc.ChanceCode) Then DicFClose.Add cc.ChanceCode, cc
                End If
            ElseIf fclose And cc.ChanceType = 2 Then
                ctimes = existChance.HoldTimeCnt + 1
                Set cc = chanceDic(cc.ChanceCode) '获得持有信息
                cc.HoldTimeCnt = cc.HoldTimeCnt + 1
                cc.UnitCost = Me.getHedgeUnitCost(cc.ChipCount, cc.BaseCost, cc.HoldTimeCnt)
                cc.Cost = cc.Cost + cc.UnitCost * cc.ChipCount
                cc.Gained = 0
                cc.CurrTimes = ctimes
                Err.Clear
                cc.Save 'cc.updateExcel 需要更新
                If Err.Number <> 0 Then
                    Dim test As String
                    test = ""
                End If
                cnt = cnt + 1
                Assetobj.CurrCash = Assetobj.CurrCash + cc.UnitCost * cc.ChipCount
                Assetobj.TotalAsset = Assetobj.TotalAsset - cc.UnitCost * cc.ChipCount
                ReDim Preserve ret(cnt)
                ret(cnt) = cc.ChanceCode & "/" & CStr(cc.BaseCost + cc.UnitCost) & "+" & Me.getRevChance(cc.ChanceCode) & "/" & CStr(cc.BaseCost)
            End If
        Else '未下注的新机 会
            
            cc.ChipCount = Me.getObjectChips(cc.ChanceCode)
            'If cc.ChipCount > 5 Then GoTo ContinueFor
            '如果资金管理方面可以允许新下注
            If (fclose = False And ((cc.ChanceType = 0 And DicFClose.Exists(cc.ChanceCode)) Or (cc.ChanceType <> 1 And cc.InputTimes = 1))) Or (fclose = True And Assetobj.AllowNewBuyForGlobel(cc.ChipCount)) Then  '补充下注不需条件
                If cc.ChanceType = 0 Then
                    '丢骰子
                    If Assetobj.AllowBuyByRnd(cc.CurrTimes, getMinTimes(cc.ChipCount), 0, getRndMaxTimes(cc.ChipCount)) Then '天要下注
                        cc.ExecDate = Format(Now(), "YYYY-MM-DD")
                        cc.ExpectCode = currExpect
                        cc.ChanceIndex = 0
                        cc.ChipCount = getObjectChips(cc.ChanceCode)
                        'cc.HoldTimeCnt = 1
                        If fclose = False Then      '如果是强平的，回退去5次,1,2组合从10开始
                            If cc.ChipCount > 2 Then
                                cc.HoldTimeCnt = cc.HoldTimeCnt * 2 / 3
                            Else
                                cc.HoldTimeCnt = 10
                            End If
                            cc.InputTimes = cc.CurrTimes '第一次下注是第几次出现机会
                        Else
                            cc.HoldTimeCnt = 1
                        End If
                        If cc.ChipCount <= 8 Then
                            cc.UnitCost = getChanceChips(cc, currExpect)
                        Else
                            cc.UnitCost = 1
                        End If
                        cc.Cost = cc.ChipCount * cc.UnitCost
                        Assetobj.CurrCash = Assetobj.CurrCash + cc.UnitCost * cc.ChipCount
                        Assetobj.TotalAsset = Assetobj.TotalAsset - cc.UnitCost * cc.ChipCount
                        cc.ChanceIndex = 0
                        cc.CreateTime = Now()
                        cc.Save 'cc.updateExcel '保存后获得indexid  需要更新
                        cnt = cnt + 1
                        ReDim Preserve ret(cnt)
                        ret(cnt) = Replace(cc.ChanceCode, "+", "/" & CStr(cc.UnitCost * gobj.SerTotal(cc.ChipCount)) & "+") & "/" & CStr(cc.UnitCost * gobj.SerTotal(cc.ChipCount))
                        chanceDic.Add cc.ChanceCode, cc
                    End If
                ElseIf cc.ChanceType = 1 Then '单次类型
                        cc.ExecDate = Format(Now(), "YYYY-MM-DD")
                        cc.ExpectCode = currExpect
                        cc.ChanceIndex = 0
                        cc.ChipCount = getObjectChips(cc.ChanceCode)
                        cc.HoldTimeCnt = 1
                        Dim checkcnt, ibeg, iend As Integer
                        checkcnt = CInt(Split(gobj.RepeatCheckCnt, "/")(0))
                        If FirstRepeatCnt = 0 Then
                            If cc.InputTimes >= checkcnt + 10 Then
                                FirstRepeatCnt = cc.InputTimes - cc.ChipCount + 1
                            Else
                                FirstRepeatCnt = cc.InputTimes - cc.ChipCount + 1
                            End If
                        End If
                        iend = cc.InputTimes - FirstRepeatCnt
                        ibeg = cc.InputTimes - FirstRepeatCnt - cc.ChipCount + 1
                        ibeg = IIf(ibeg < 1, 1, ibeg)
                        ibeg = CInt(gobj.UnitChipArray(1)(ibeg))
                        iend = CInt(gobj.UnitChipArray(1)(iend))
                        Dim UCost As Integer
                        UCost = IIf(cc.ChipCount < 4, (ibeg + iend) / 2, iend)
                        If UCost <= 0 Then UCost = 1
                        cc.UnitCost = CStr(UCost)
                        cc.Cost = cc.ChipCount * cc.UnitCost
                        Assetobj.CurrCash = Assetobj.CurrCash + cc.UnitCost * cc.ChipCount
                        Assetobj.TotalAsset = Assetobj.TotalAsset - cc.UnitCost * cc.ChipCount
                        cc.CreateTime = Now()
                        cc.Save
                        cnt = cnt + 1
                        ReDim Preserve ret(cnt)
                        ret(cnt) = Replace(cc.ChanceCode, "+", "/" & CStr(cc.UnitCost * gobj.SerTotal(1)) & "+") & "/" & CStr(cc.UnitCost * gobj.SerTotal(1))
                        chanceDic.Add cc.ChanceCode, cc
                ElseIf cc.ChanceType = 2 Then '对冲类型
                        cc.ExecDate = Format(Now(), "YYYY-MM-DD")
                        cc.ExpectCode = currExpect
                        cc.ChanceIndex = 0
                        cc.ChipCount = getObjectChips(cc.ChanceCode)
                        cc.HoldTimeCnt = 1
                        cc.UnitCost = Me.getHedgeUnitCost(cc.ChipCount, cc.BaseCost, 1)
                        cc.Cost = cc.ChipCount * cc.UnitCost
                        Assetobj.CurrCash = Assetobj.CurrCash + cc.UnitCost * cc.ChipCount
                        Assetobj.TotalAsset = Assetobj.TotalAsset - cc.UnitCost * cc.ChipCount
                        cc.CreateTime = Now()
                        cc.Save
                        cnt = cnt + 1
                        ReDim Preserve ret(cnt)
                        ret(cnt) = cc.ChanceCode & "/" & CStr(cc.BaseCost + cc.UnitCost) & "+" & Me.getRevChance(cc.ChanceCode) & "/" & CStr(cc.BaseCost)
                        chanceDic.Add cc.ChanceCode, cc
                End If
            End If
        End If
ContinueFor:
    Next
    getChips = Trim(Join(ret, " "))
End Function

Function getHedgeUnitCost(chips As Integer, baseUnitCost As Integer, HoldTimes As Integer)
    If HoldTimes = 0 Then
        getHedgeUnitCost = 0
        Exit Function
    Else
        Dim val As Single
        val = (getHedgeSum(chips, baseUnitCost, HoldTimes - 1) * chips - baseUnitCost * (gobj.Odds - 10) * HoldTimes) / (gobj.Odds - chips)
        getHedgeUnitCost = Fix(val) + 1
        Exit Function
    End If
End Function

Public Function getHedgeSum(chips As Integer, baseUnitCost As Integer, HoldTimes As Integer) As Integer
    Dim i  As Integer
    Dim ret As Integer
    ret = 0
    For i = 1 To HoldTimes
        Dim currVal As Integer
        currVal = getHedgeUnitCost(chips, baseUnitCost, i)
        ret = ret + currVal
    Next
    getHedgeSum = ret
End Function

Function ToSerial(str As String, MinChips As Integer)
    str = Trim(str)
    Dim arr() As String
    arr = Split(str, " ")
    Dim dic As New Dictionary
    Set dic = New Dictionary
    Dim dicint  As New Dictionary
    Set dicint = New Dictionary
    Dim i, j As Integer
    For i = 0 To UBound(arr)
        Dim sarr() As String
        sarr = Split(arr(i), "+")
        For j = 0 To UBound(sarr)
            Dim car, serno, cc As String
            Dim cnt As Integer
            Dim carr() As String
            carr = Split(sarr(j), "/")
            car = carr(0)
            serno = carr(1)
            cnt = CInt(carr(2))
            For k = 1 To Len(car)
                
                For m = 1 To Len(serno)
                    Dim carc, serc As Integer
                    carc = IIf(Mid(car, k, 1) = "0", 10, CInt(Mid(car, k, 1)))
                    serc = IIf(Mid(serno, m, 1) = "0", 10, CInt(Mid(serno, m, 1)))
                    cc = Mid(car, k, 1) & "/" & Mid(serno, m, 1)
                    If dic.Exists(cc) Then
                        dic(cc) = dic(cc) + cnt
                    Else
                        dic.Add cc, cnt
                        dicint.Add 100 * carc + serc, cc
                    End If
                Next
            Next
        Next
    Next
    Dim ret() As String
    Dim n As Integer

    ReDim retint(dic.Count)
    n = 0
    For Each k In dicint.Keys
        n = n + 1
        retint(n) = k
    Next
    
    Dim minval As Long
    For i = 1 To UBound(retint)
        minval = retint(i)
        For j = i + 1 To UBound(retint)
            If retint(j) < minval Then
                minval = retint(i)
                retint(i) = retint(j)
                retint(j) = minval
                minval = retint(i)
            End If
        Next
    Next
    cnt = 1
    For i = 1 To UBound(retint)
        ival = dic(dicint(retint(i)))
        If ival >= MinChips Then
            ReDim Preserve ret(dic.Count)
            ret(cnt) = dicint(retint(i)) & "/" & dic(dicint(retint(i)))
            cnt = cnt + 1
        End If
    Next
    ToSerial = Join(ret, " ")
End Function

Function getMinTimes(Cols As Integer)
    If Cols = 1 Then
        getMinTimes = iOneCodeMinCheckTimes + IIf(Hour(Now()) > 12, 3, 0)
    Else
        getMinTimes = iMinCheckTimes + 3 * (5 - Cols) + IIf(Hour(Now()) > 12, 1, 0)
    End If
    'If cols = 3 And getMinTimes <= 19 Then getMinTimes = 20
End Function

Function getRndMaxTimes(Cols As Integer)
    If Cols = 1 Then
        getRndMaxTimes = IIf(Hour(Now()) > 17, 5, IIf(Hour(Now()) < 13, 1, 3))
    Else
        If Cols = 3 Then
            getRndMaxTimes = IIf(Hour(Now()) > 19, 2, IIf(Hour(Now()) < 13, 0, 1))
        Else
            getRndMaxTimes = IIf(Hour(Now()) > 17, 0, IIf(Hour(Now()) < 13, 0, 0))
        End If
    End If
End Function


Function getChanceChips(chance As ChanceClass, currExpect As String) As Integer

'''    'Dim sheet As Worksheet
'''    If chance.ChipCount = 3 Then Set sheet = Sheet6
'''    If chance.ChipCount = 4 Then Set sheet = Sheet7
'''    If chance.ChipCount = 5 Then Set sheet = Sheet8
'''    If chance.ChipCount = 6 Then Set sheet = Sheet13
'''    If chance.ChipCount = 8 Then Set sheet = Sheet14
'''    If chance.ChipCount = 1 Then Set sheet = Sheet5
'''
    'expectcnt = CLng(CurrExpect) - CLng(chance.ExpectCode)
    expectcnt = chance.HoldTimeCnt  ' CLng(CurrExpect) - CLng(chance.ExpectCode)
    arr = gobj.UnitChipArray(chance.ChipCount)
    getChanceChips = arr(expectcnt - 1)
End Function

Public Function SaveDayRecs()
    Dim i As Integer
    Dim strDate As String
    Dim tmpdate As String
    With Sheet2
        i = 0
        strDate = .cells(1, 8 + i).Text
        tmpdate = strDate
        While Len(tmpdate) > 0
            i = i + 1
            tmpdate = .cells(1, 8 + i).Text
            If Len(tmpdate) > 0 Then strDate = tmpdate
        Wend
        If Len(strDate) > 0 Then
            If day(strDate) = day(DateAdd("D", -1, Now())) Then
                Exit Function
            End If
        End If
         .cells(1, 8 + i) = Format(DateAdd("D", -1, Now()), "YYYY-MM-DD")
        ls = .cells(2, 7)
        yl = .cells(3, 7)
        ky = .cells(4, 7)
        .cells(2, 7) = 0
        .cells(3, 7) = 0
        .cells(4, 7) = .cells(1, 7)
        .cells(2, 8 + i) = ls
        .cells(3, 8 + i) = yl
        .cells(4, 8 + i) = ky
    End With
    SaveDayRecs = True
End Function

Public Function getAllSerialDataDic(ByRef ret As Dictionary)
    On Error Resume Next
    Set ret = New Dictionary
    Dim firstNo As String
    Dim datas() As ExpectData
    Dim ed As New ExpectData
    ed.TableToArray ed.GetHistoryData(Me.ValidOldestHistoryExpect), datas
    If UBound(datas) <= 0 Then Exit Function
    Dim strno As String
    Dim i As Long
    i = 1
    'Dim ed As ExpectData
    Set ed = datas(1)
    strno = ed.Expect
    ret.Add strno, New ExpectData
    ret(strno).OpenCode = ed.OpenCode
    ret(strno).Expect = ed.Expect
    ret(strno).OpenTime = ed.OpenTime
    i = i + 1
    Set ed = Nothing
    Set ed = datas(i)
    While i <= UBound(datas) And CLng(ret(strno).ExpectIndex) - 1 = ed.ExpectIndex
        Set ed = Nothing
        Set ed = datas(i)
        strno = ed.Expect
        If ret.Exists(strno) = False Then
           ret.Add strno, New ExpectData
        Else
            'MsgBox strno
            test = ""
        End If
        ret(strno).OpenCode = ed.OpenCode
        ret(strno).Expect = ed.Expect
        ret(strno).OpenTime = ed.OpenTime
        i = i + 1
        Set ed = Nothing
        Set ed = datas(i)
        ValidOldestHistoryExpect = strno
''        Exit Function
        'Sheet16.Cells(1, 12) = firstNo
''        Exit Function
''        While i > 1
''
''            If CLng(strno) - 1 <> .Cells(i, 1) Then
''                .Range(Replace("XX:XX", "XX", i)).Interior.Color = RGB(255, 255, 0)
''                'MsgBox i
''            End If
''            strno = .Cells(i, 1).Text
''            i = i - 1
         If i > UBound(datas) Then Exit Function
        Wend
        
    'End With
    
End Function

Public Function getSerialCars(ByVal dic As Dictionary, ByVal expertNo As String, n As Integer, FromN As Integer, ToN As Integer) As String()
    Dim ret() As String
    ReDim ret(0)
    startno = CStr(CLng(expertNo) - n + 1)
    Dim cnt As Integer
    If dic.Exists(expertNo) = False Or dic.Exists(startno) = False Then
        getSerialCars = ret
        Exit Function
    End If
    For i = FromN To ToN
        Dim ed As ExpectData
        Set ed = dic(expertNo)
        Dim carno As String
        carno = Split(ed.OpenCode, ",")(i - 1)
        If carno = "**" Then GoTo NoSeeNo
        For j = 1 To n - 1
            preno = CStr(CLng(expertNo) - j)
            If carno <> Split(dic(preno).OpenCode, ",")(i - 1) Then GoTo NoSeeNo
        Next
        cnt = cnt + 1
        ReDim Preserve ret(cnt)
        ret(cnt) = CStr(i) & "/" & Replace(Int(carno), "10", "0")
NoSeeNo:
    Next
    getSerialCars = ret
End Function

Public Sub Research(ByRef retChance() As ChanceClass, dg As MSFlexGrid, ByRef currChance As String, ByRef holdingCnt As Integer, ByRef HoldingExpectCnt As Integer, ByRef lastDate As String)
    On Error Resume Next
    'dg.Clear
    Dim DataDic As Dictionary
    'getAllSerialDataDic DataDic
    Dim cars() As String
    Dim allcnt As Long
    Dim ContinueSameCnt As Long
    Dim strKey As Object
    Dim LastSameNo As String
    Dim lastChance As String
    'Dim currChance As String
    Dim cnt As Long
    Dim LastChanceCnt As Long
    'Dim lastDate As String
    Dim lastSameCnt As Integer
    Dim expertCnt As Long
    Dim lastSameExpert As Long
    Dim lastSameCarCnt As Long
    'Sheet16.Range("N2:T30000").Clear
    'ContinueSameCnt = 1
    Set DataDic = Me.DicAllUseData
    Dim eds() As ExpectData
    Dim key As Variant
    For Each key In DataDic.Keys
        expertCnt = expertCnt + 1
        ReDim Preserve eds(expertCnt)
        Set eds(expertCnt) = DataDic(key)
    Next
    expertCnt = 0
    Dim Ei As Long
    For Ei = UBound(eds) To 1 Step -1
        expertCnt = expertCnt + 1
        key = eds(Ei).Expect
        cars = getSerialCars(DataDic, key, gobj.SingleCarRepeatCnt, gobj.ResearchStartCol, 10)
        If UBound(cars) > 0 Then cnt = cnt + 1
        If key = Me.currExpect Then '如果最后一期
            If UBound(cars) > 0 Then currChance = Replace(Join(cars, " "), "10", "0")
        End If
        allcnt = allcnt + UBound(cars)
        NextNo = CStr(CLng(key) + 1)
        If DataDic.Exists(NextNo) Then
            For j = 1 To UBound(cars)
               arr = Split(cars(j), "/")
               POS = CInt(arr(0))
               carno = Right("00" & Replace(arr(1), "0", "10"), 2)
               Dim ed As ExpectData
               Set ed = DataDic(NextNo)
               If Split(ed.OpenCode, ",")(POS - 1) = carno Then
                    If Len(Trim(LastSameNo)) = 0 Then
                        LastSameNo = NextNo
                        lastChance = cars(j)
                        LastChanceCnt = cnt
                        lastDate = ed.OpenTime
                    End If
                    ContinueSameCnt = ContinueSameCnt + 1
'''                    Sheet16.Range(Replace("N$:N$", "$", ContinueSameCnt + 1)) = NextNo
'''                    Sheet16.Range(Replace("O$:O$", "$", ContinueSameCnt + 1)) = ed.OpenTime
'''                    Sheet16.Range(Replace("P$:P$", "$", ContinueSameCnt + 1)) = Replace(cars(j), "/", "->")
'''                    Sheet16.Range(Replace("Q$:Q$", "$", ContinueSameCnt + 1)) = cnt - lastSameCnt
'''                    Sheet16.Range(Replace("R$:R$", "$", ContinueSameCnt + 1)) = expertCnt - lastSameExpert
'''                    Sheet16.Range(Replace("S$:S$", "$", ContinueSameCnt + 1)) = Split(cars(j), "/")(0)
'''                    Sheet16.Range(Replace("T$:T$", "$", ContinueSameCnt + 1)) = Split(cars(j), "/")(1)
                    'dg.TextMatrix(n+1,ContinueSameCnt + 1)
                    dg.Rows = ContinueSameCnt + 1
                    dg.TextMatrix(ContinueSameCnt, 0) = ContinueSameCnt
                    dg.TextMatrix(ContinueSameCnt, 1) = NextNo
                    dg.TextMatrix(ContinueSameCnt, 2) = ed.OpenTime
                    dg.TextMatrix(ContinueSameCnt, 3) = Replace(cars(j), "/", "->")
                    dg.TextMatrix(ContinueSameCnt, 4) = allcnt - lastSameCarCnt
                    dg.TextMatrix(ContinueSameCnt, 5) = cnt - lastSameCnt
                    dg.TextMatrix(ContinueSameCnt, 6) = expertCnt - lastSameExpert
                    dg.TextMatrix(ContinueSameCnt, 7) = Split(cars(j), "/")(0)
                    dg.TextMatrix(ContinueSameCnt, 8) = Split(cars(j), "/")(1)
                    lastSameExpert = expertCnt
                    lastSameCnt = cnt
                    lastSameCarCnt = allcnt
                   ' MsgBox Replace("[XX]", "XX", nextno) & cars(J) & "->" & ed.opencode
    
               End If
            Next
        End If
    Next
    holdingCnt = cnt
    HoldingExpectCnt = currExpect - LastSameNo
''''    Sheet16.Cells(1, 11) = allcnt
''''    Sheet16.Cells(2, 11) = ContinueSameCnt
''''    Sheet16.Cells(4, 11) = LastSameNo
''''    Sheet16.Cells(4, 12) = lastDate
''''    Sheet16.Cells(5, 11) = Sheet16.Cells(1, 6) - LastSameNo
''''    Sheet16.Cells(5, 12) = DataDic.Count / ContinueSameCnt
''''    Sheet16.Cells(6, 11) = lastChance
''''    Sheet16.Cells(6, 12) = LastChanceCnt
''''    Sheet16.Cells(7, 11) = currChance
''''
''''    Sheet1.Range("Q2:Q2") = currChance
''''    Sheet1.Range("P2:T2").Interior.Color = RGB(255, 255, 255)
''''    Sheet1.Range("P2:P2") = ""
''''    If Len(Trim(currChance)) > 0 Then
''''        'Sheet1.Range("P2:T2").Interior.Color = RGB(255, 255, 0)
''''        Items = Split(currChance, " ")
''''        For I = 1 To UBound(Items)
''''           chance = Items(I)
''''           pc = Split(chance, "/")
''''           strPos = Replace(pc(0), "10", 0)
''''           strCar = Replace("1234567890", pc(1), "")
''''           Items(I) = Replace(strPos & "/" & strCar, " ", "")
''''        Next
''''        'Sheet1.Range("P2:P2") = Join(Items, " ")
''''    End If
''''    Sheet1.Range("R2:R2") = Sheet1.Cells(2, 2) - LastSameNo
''''    Sheet1.Range("S2:S2") = DataDic.Count / ContinueSameCnt
''''    Sheet1.Range("T2:T2") = LastChanceCnt
    Dim dgc As New DataGridClass
    Set dgc.grid = dg
    dgc.AdjustColWidth dg.Parent, dg
End Sub

Public Sub InitHistoryData()
    Dim exparr() As ExpectData
    Dim ed As New ExpectData
    Set ed.db.gobj = Me.gobj
    ed.TableToArray ed.GetHistoryData(Me.ValidOldestHistoryExpect, "100"), exparr
    getAllDataDic exparr, Me.DicAllHistoryData
End Sub

Sub GetHistoryData(ByRef AccessHistoryPages As Long)
    On Error Resume Next
    Dim hdc As New HistoryDataClass
    Dim exparr() As ExpectData
    Dim i As Long
    'Dim allDataDic As New Dictionary
    ReadHistoryData exparr
    gobj.LogObj.Log "获取历史数据字典"
    Set DicAllHistoryData = getAllDataDic(exparr, Me.DicAllHistoryData)
    i = AccessHistoryPages
    grp = i
    'Sheet16.Range("b4:b4").Interior.Color = RGB(255, 255, 2550)
    While i > 0
        gobj.LogObj.Log "获取历史数据页数：" & CStr(i)
        hdc.GetHistoryData i, exparr
        If UBound(exparr) <= 0 Then Exit Sub
        NewestHistoryExpect = exparr(1).Expect
        If NewestHistoryExpect = CurrHistoryExpect Then Exit Sub
        If UBound(exparr) > 0 Then
            gobj.LogObj.Log "保存新的历史数据"
            WriteAllData exparr, DicAllHistoryData
            'Set allDataDic = getAllDataDic(exparr, allDataDic)
            CurrHistoryExpect = NewestHistoryExpect
        Else
            'MsgBox "无法获取数据！"
            'Sheet16.Range("b4:b4").Interior.Color = RGB(255, 0, 0)
            test = ""
            Exit Sub
        End If
        i = i - 1
        'Sheet16.Cells(4, 2) = I
        AccessHistoryPages = IIf(i > 1, i, 10)
        ReDim exparr(0)
        If grp - i = 10 Then Exit Sub
    Wend
 
End Sub


Private Sub Class_Initialize()
    Set Me.DicAllHistoryData = New Dictionary
    Set Me.DicAllUseData = New Dictionary
End Sub
