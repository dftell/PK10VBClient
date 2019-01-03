VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChanceClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ExpectCode As String
Public ChanceCode As String '
Public UnitCost  As String
Public ChipCount As Integer
Public ChanceIndex As Long
Public ExecDate As String
Public IsEnd As Integer
Public HoldTimeCnt As Integer
Public Cost As Long
Public Gained As Single
Public InputTimes As Integer '下注时出现的次数
Public CurrTimes As Integer '不记录
Public Profit As Single
Public CreateTime As Date
Public UpdateTime As Date
Public BaseCost As Long
Public Remark As String
Public db As New DBClass
Public ChanceType As Integer '0,持续类型，1，一次性类型
Public Sub TableToArray(dt As DataTable, ByRef arr() As ChanceClass)
    If dt Is Nothing Then Exit Sub
    ReDim arr(dt.RowCount)
    Dim i As Long
    For i = 1 To dt.RowCount
        Dim dic As Dictionary
        Set dic = dt.Rows(i)
        If dic Is Nothing Then
            Exit Sub
        End If
        ReDim Preserve arr(i)
        Set arr(i) = New ChanceClass
        arr(i).ChanceCode = dic("ChanceCode")
        arr(i).ExpectCode = dic("ExpectCode")
        arr(i).ChanceType = IIf(IsNull(dic("ChanceType")), 0, dic("ChanceType"))
        arr(i).UnitCost = dic("UnitCost")
        arr(i).ChipCount = dic("ChipCount")
        If (dic.Exists("ChanceIndex")) Then
            arr(i).ChanceIndex = IIf(IsNull(dic("ChanceIndex")), 0, dic("ChanceIndex"))
        End If
        arr(i).ExecDate = dic("ExecDate")
        arr(i).IsEnd = dic("IsEnd")
        arr(i).HoldTimeCnt = dic("HoldTimeCnt")
        arr(i).Cost = dic("Cost")
        arr(i).Gained = dic("Gained")
        arr(i).InputTimes = dic("InputTimes")
        arr(i).CurrTimes = dic("CurrTimes")
        arr(i).Profit = dic("Profit")
        arr(i).CreateTime = dic("CreateTime")
        arr(i).UpdateTime = dic("UpdateTime")
        arr(i).BaseCost = dic("BaseCost")
        arr(i).Remark = dic("Remark")
    Next
    Exit Sub
End Sub

Public Function ArrayToTable(arr() As ChanceClass, Optional excludeCols As String = "") As DataTable
    Dim ret As New DataTable
    ret.FillByArray "ChanceIndex,ExpectCode,ChanceCode,UnitCost,ChipCount,ExecDate,IsEnd,HoldTimeCnt,Cost,Gained,InputTimes,CurrTimes,Profit,CreateTime,UpdateTime,ChanceType,BaseCost,Remark", arr, 1, excludeCols
    Set ArrayToTable = ret
End Function

'''''Public Sub updateExcel()
'''''    Dim needcreate As Boolean
'''''    If ChanceIndex = 0 Then
'''''        ChanceIndex = Sheet12.UsedRange.Rows.Count
'''''        needcreate = True
'''''    End If
'''''    'ChanceIndex ExpectCode  ChanceCode  ChipCount   UnitCost
'''''    If ExecDate = "" Then ExecDate = Format(Now(), "YYYY-MM-DD")
'''''    If ChipCount = 0 Then
'''''        codes = Split(ChanceCode, "/")
'''''        ChipCount = Len(codes(0)) * Len(codes(1))
'''''    End If
'''''    If Cost = 0 Then
'''''        Cost = Me.UnitCost * Me.ChipCount
'''''    End If
'''''    If Gained = 0 Then
'''''        Gained = 0
'''''    End If
'''''    Dim rowid As Long
'''''    With Sheet12
'''''        rowid = ChanceIndex + 1
'''''        .Cells(rowid, 1) = ChanceIndex
'''''        .Cells(rowid, 2) = ExpectCode
'''''        .Cells(rowid, 3) = ChanceCode
'''''        .Cells(rowid, 4) = ChipCount
'''''        .Cells(rowid, 5) = UnitCost
'''''        .Cells(rowid, 6) = ExecDate
'''''        .Cells(rowid, 7) = IIf(IsEnd, 1, 0)
'''''        .Cells(rowid, 8) = HoldTimeCnt
'''''        .Cells(rowid, 9) = Cost
'''''        .Cells(rowid, 10) = Gained
'''''        .Cells(rowid, 11) = InputTimes
'''''        .Cells(rowid, 12) = Profit
'''''        .Cells(rowid, 13) = CurrTimes
'''''        .Cells(rowid, 15) = Now()
'''''        If needcreate Then .Cells(rowid, 14) = Now()
'''''    End With
'''''End Sub

Public Function Save() As Boolean
    Dim dt As DataTable
    Dim arr(1) As ChanceClass
    Me.UpdateTime = Now()
    Set arr(1) = Me
    Set dt = ArrayToTable(arr)
    Dim sql As String
    sql = "select * from ChanceTable where ChanceIndex=" & ChanceIndex
    If Me.ChanceIndex <= 0 Then
        sql = "select * from ChanceTable where ChanceIndex < -9999"
    End If
    db.UpdateDataBuySql sql, dt, "ChanceIndex"
End Function

Public Function getDayData(day As Date) As DataTable
    Dim sql As String
    Dim ret As DataTable
    sql = Replace("select * from ChanceTable where [ExecDate]>='XX'  order by createtime desc", "XX", day)
    Set ret = db.getDataBySql(sql)
    Set getDayData = ret
End Function

Public Function getSpecDayData(day As Date) As DataTable
    Dim sql As String
    Dim ret As DataTable
    sql = Replace("select * from ChanceTable where [ExecDate]='XX'  order by createtime desc", "XX", day)
    Set ret = db.getDataBySql(sql)
    Set getSpecDayData = ret
End Function

Public Function getDayHoldingData(day As Date) As DataTable
    Dim sql As String
    Dim ret As DataTable
    sql = Replace("select * from ChanceTable where [ExecDate]>='XX' and [IsEnd]=0  order by createtime desc", "XX", day)
    Set ret = db.getDataBySql(sql)
    Set getDayHoldingData = ret
End Function

Public Function getChanceById(index As Long) As DataTable
    Dim sql As String
    Dim ret As DataTable
    sql = Replace("select * from ChanceTable where [ChanceIndex]=XX order by Expect desc", "XX", index)
    Set ret = db.getDataBySql(sql)
    Set getChanceById = ret
End Function

Public Function getSummary() As DataTable
    Dim sql As String
    Dim ret As DataTable
    sql = Replace("select ExecDate,sum(Cost) as '流水',sum(Gained) as '盈利',sum(Profit) as '净利',sum(holdtimecnt) as '合计次数',max(cost) as '最大成本',100*sum(profit)/sum(cost) as '日收益率',sum(Profit)/max(cost) as RRR,max(holdtimeCnt) as '最多持有次数' from ChanceTable group by execDate order by execDate desc", "XX", index)
    Set ret = db.getDataBySql(sql)
    Set getSummary = ret
End Function

Public Sub DeleteDayData(day As Date)
    Dim sql As String
    sql = Replace("delete from ChanceTable where [ExecDate]='XX'", "XX", day)
    db.ExcelSql (sql)
End Sub

Public Sub DeleteChance(index As String)
    Dim sql As String
    sql = Replace("delete from ChanceTable where [ChanceIndex]=XX", "XX", index)
    db.ExcelSql (sql)
End Sub

Public Function getChance(index As Long) As ChanceClass
    Dim dt As DataTable
    Set dt = getChanceById(index)
    If Not dt Or dt.RowCount = 0 Then Exit Function
    Dim arr() As ChanceClass
    Call TableToArray(dt, arr)
    Set getChance = arr(1)
    Exit Function
    Dim ret As New ChanceClass
    Dim i As Long
    i = index + 1
    Exit Function
''''    'With  Sheet12
''''        If i > .UsedRange.Rows.Count Then
''''            Set getChance = Nothing
''''            Exit Function
''''        End If
''''        ret.ChanceIndex = CLng(.Cells(i, 1))
''''        ret.ExpectCode = .Cells(i, 2)
''''        ret.ChanceCode = .Cells(i, 3)
''''        ret.ChipCount = .CInt(Cells(i, 4))
''''        ret.UnitCost = CInt(.Cells(i, 5))
''''        ret.ExecDate = .Cells(i, 6).Text
''''        ret.IsEnd = IIf(.Cells(i, 7).Text = "1", True, False)
''''        ret.HoldTimeCnt = CInt(.Cells(i, 8).Text)
''''        ret.Cost = CInt(.Cells(i, 9).Text)
''''        ret.Gained = CInt(.Cells(i, 10).Text)
''''        ret.InputTimes = CInt(.Cells(i, 11).Text)
''''        ret.CurrTimes = .Cells(i, 13).Text
''''    'End With
    Set getChance = ret
End Function

Public Sub ReadCurrDateHoldingChances(ByRef ret() As ChanceClass)
    Dim dt As DataTable
    Set dt = Me.getDayHoldingData(DateAdd("D", -1, Date)) '为支持最后一期仍未中奖，需要取前一天的交易记录
    Me.TableToArray dt, ret
    Exit Sub
    
''''''''''
''''''''''
''''''''''    On Error Resume Next
''''''''''    Dim lastCount As Long
''''''''''    lastCount = Sheet12.UsedRange.Rows.Count
''''''''''    Dim I As Long
''''''''''    Dim cnt As Long
''''''''''    With Sheet12
''''''''''        For I = lastCount To 2 Step -1
''''''''''            strDate = .Cells(I, 6).Text
''''''''''            strEnd = .Cells(I, 7).Text
''''''''''            If DateDiff("D", Now(), CDate(strDate)) > 0 Then
''''''''''                Exit Sub
''''''''''            End If
''''''''''            If strEnd = "1" Then
''''''''''            Else
''''''''''                cnt = cnt + 1
''''''''''                ReDim Preserve ret(cnt)
''''''''''                Set ret(cnt) = New ChanceClass
''''''''''                ret(cnt).ChanceIndex = CLng(.Cells(I, 1).Value)
''''''''''                ret(cnt).ExpectCode = .Cells(I, 2).Value
''''''''''                ret(cnt).ChanceCode = .Cells(I, 3).Value
''''''''''                ret(cnt).ChipCount = CInt(.Cells(I, 4).Value)
''''''''''                ret(cnt).UnitCost = CInt(.Cells(I, 5).Value)
''''''''''                ret(cnt).ExecDate = .Cells(I, 6).Text
''''''''''                ret(cnt).IsEnd = IIf(.Cells(I, 7).Value = "1", True, False)
''''''''''                ret(cnt).HoldTimeCnt = CInt(.Cells(I, 8).Value)
''''''''''                ret(cnt).Cost = CInt(.Cells(I, 9).Value)
''''''''''                ret(cnt).Gained = CInt(.Cells(I, 10).Value)
''''''''''                ret(cnt).InputTimes = CInt(.Cells(I, 11).Value)
''''''''''                ret(cnt).CurrTimes = CInt(.Cells(I, 13).Text)
''''''''''            End If
''''''''''        Next
''''''''''    End With
End Sub

Public Function getHoldChancesDic(Chances() As ChanceClass) As Dictionary
    Dim ret As New Dictionary
    For i = 1 To UBound(Chances)
        If Not ret.Exists(Chances(i).ChanceCode) Then
            ret.Add Chances(i).ChanceCode, Chances(i)
        End If
    Next
    Set getHoldChancesDic = ret
End Function


Private Sub Class_Initialize()
    ChanceIndex = 1
End Sub

Public Function Copy() As ChanceClass
    On Error Resume Next
    Dim ret As New ChanceClass
    Dim tlbapp As New TLIApplication
    
    Dim tlbinfo As InterfaceInfo
    Set tlbinfo = tlbapp.InterfaceInfoFromObject(Me)
    Dim meminfo As MemberInfo
    For Each key In tlbinfo.Members
        Set meminfo = key
        Dim res As Variant
        Err.Clear
        res = CallByName(Me, meminfo.Name, VbGet)
        If Err.Number <> 0 Then GoTo ContinueFor
            If Not IsEmpty(res) And Not IsNull(res) Then
                CallByName ret, meminfo.Name, VbLet, res
            End If
ContinueFor:
    Next
    Set Copy = ret
End Function
