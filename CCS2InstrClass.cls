VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCS2InstrClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ruleId As String
Public betNum As Integer    '数量
Public itemTimes As String
Public selNums As String
Public jsOdds As Single
Public priceMode As String
Const cRuleId_S = "8140101" '前5位定胆
Const cRuleId_B = "8140102" '后5位定胆
Public cJsOdds As String
Const strJsonModel = "{""ruleId"":""{ruleId}"",""betNum"":""{betNum}"",""itemTimes"":{itemTimes},""selNums"":""{selNums}"",""jsOdds"":""{jsOdds}"",""priceMode"":{priceMode}}"
Public Function InstrToJsonString(ccs As String, Optional unit As Integer = 2) As String 'unit为单位 0，元；1，角；2，分，3，其他
    ccs = Replace(ccs, "  ", "")
    Dim cm As New ClassModel
    ccs = Trim(ccs)
    ccs = Replace(ccs, "+", " ") '去掉所有的加号，替换为空格
    'rowcnt = UBound(Split(ccs, " ")) + 1
    If Len(ccs) = 0 Then
        InstrToJsonString = ""
        Exit Function
    End If
    ccsarr = Split(ccs, " ")
    Dim i As Integer
    Dim j As Integer
    Dim c As Integer
    Dim InsArr() As CCS2InstrClass
    Dim ArrCnt As Integer
    Dim unitVal As String
    unitVal = getUnitValue(unit)
    ArrCnt = 0
    For i = 0 To UBound(ccsarr)
        Dim cc As String, ccNos As String, ccCars As String, ccUnitCost As Long, ccArr() As String
        cc = Trim(ccsarr(i))
        ccArr = Split(cc, "/")
        If Len(cc) = 0 Then GoTo EndFor
        ccNos = ccArr(0)
        Dim ccOrgCars As String
        ccOrgCars = ccArr(1)
        ccCars = toStdCarFmt(ccArr(1))
        ccUnitCost = ccArr(2)
        Dim sarr() As String
        Dim bArr() As String
        Dim sArrCnt As Integer, bArrCnt As Integer
        sArrCnt = 0
        bArrCnt = 0
        ReDim sarr(4)
        ReDim bArr(4)
        For j = 1 To Len(ccNos)
            Dim strsNo As String
            Dim iNo As Integer
            strsNo = Mid(ccNos, j, 1)
            If strsNo = "0" Then strsNo = "10"
            iNo = CInt(strsNo)
            If iNo <= 5 Then
                sarr(iNo - 1) = ccCars
                sArrCnt = sArrCnt + 1
            Else
                bArr(iNo - 5 - 1) = ccCars
                bArrCnt = bArrCnt + 1
            End If
        Next
        If sArrCnt > 0 Then
            ReDim Preserve InsArr(ArrCnt)
            Set InsArr(ArrCnt) = New CCS2InstrClass
            InsArr(ArrCnt).ruleId = cRuleId_S
            InsArr(ArrCnt).betNum = Len(ccOrgCars) * sArrCnt '车数*前后区间所有车号数量
            InsArr(ArrCnt).itemTimes = Format(unitVal * CLng(ccUnitCost), "0.00")
            InsArr(ArrCnt).selNums = Join(sarr, ",")
            InsArr(ArrCnt).jsOdds = cJsOdds
            InsArr(ArrCnt).priceMode = unit
            ArrCnt = ArrCnt + 1
        End If
        If bArrCnt > 0 Then
            ReDim Preserve InsArr(ArrCnt)
            Set InsArr(ArrCnt) = New CCS2InstrClass
            InsArr(ArrCnt).ruleId = cRuleId_B
            InsArr(ArrCnt).betNum = Len(ccOrgCars) * bArrCnt '车数*前后区间所有车号数量
            InsArr(ArrCnt).itemTimes = Format(unitVal * CLng(ccUnitCost), "0.00")
            InsArr(ArrCnt).selNums = Join(bArr, ",")
            InsArr(ArrCnt).jsOdds = cJsOdds
            InsArr(ArrCnt).priceMode = unit
            ArrCnt = ArrCnt + 1
        End If
EndFor:
    Next
    If UBound(InsArr) >= 0 Then
        InstrToJsonString = InstToJsonString(InsArr)
    Else
        InstrToJsonString = ""
    End If
End Function

Function InstToJsonString(Insts() As CCS2InstrClass) As String
    If UBound(Insts) < 0 Then
        InstToJsonString = ""
        Exit Function
    End If
    Dim i As Integer
    Dim strArr() As String
    ReDim strArr(IIf(UBound(Insts) = 0, 1, UBound(Insts)))
    For i = 0 To UBound(Insts)
        strInst = Replace(strJsonModel, "{ruleId}", Insts(i).ruleId)
        strInst = Replace(strInst, "{betNum}", Insts(i).betNum)
        strInst = Replace(strInst, "{itemTimes}", Insts(i).itemTimes)
        strInst = Replace(strInst, "{selNums}", Insts(i).selNums)
        strInst = Replace(strInst, "{jsOdds}", Insts(i).jsOdds)
        strInst = Replace(strInst, "{priceMode}", Insts(i).priceMode)
        strArr(i) = strInst
    Next
    If UBound(Insts) = 0 Then
        InstToJsonString = "[" & strArr(0) & "]"
    Else
        InstToJsonString = "[" & Join(strArr, ",") & "]"
    End If
End Function

Function getUnitValue(unit As Integer) As Single
    Dim str As String
    getUnitValue = 0.1 ^ unit
End Function

Function toStdCarFmt(cars As String) As String
    cars = Trim(cars)
    Dim CarArr() As String
    ReDim CarArr(Len(cars))
    Dim i As Integer
    For i = 1 To Len(cars)
        Dim strCar As String
        strCar = Mid(cars, i, 1)
        If strCar = "0" Then strCar = "10"
        CarArr(i) = Right("0" & strCar, 2)
    Next
    toStdCarFmt = Trim(Join(CarArr, " "))
End Function

'''for(h=0;h<nt.length;h++)
'''{
'''    var u=$(nt[h]),
'''    g=u.find(".betting-multiple").val(), 倍数
'''    tt=t(g);
'''    if(tt!="")
'''        return u.find(".betting-multiple").focus(),alert(tt),!1;
'''    a+=" "+u.find(".bof-tb-mode").text(); '玩法 8110101/8110102
'''    var it=u.find(".bof-tb-method>i"), '模式
'''    ut=it.attr("pricevalue"), '
'''    rt=$.trim(it.html()),  角 0.1 元 1 分0.01
'''    p={ruleId:u.attr("ruleid"),
'''    betNum:u.find(".bof-tb-count>em").text(), 注数
'''    itemTimes:g*ut, , 倍数*单位
'''    selNums:u.find(".bof-tb-number").attr("allnums"), 选号 ，，01 02，，，形式
'''    jsOdds:u.attr("jsodds"), 赔率 9.75
'''    priceMode:GetBetPriceMode(rt)}; function GetBetPriceMode(n){return n=="元"?0:n=="角"?1:n=="分"?2:3}
'''    l.push(p),
'''    gl.isK3||(b=GetDataDetailById(p.ruleId,gl.rules,"RuleId"),p.betNum*100/b.Modulus<=2&&(s+=","+(h+1)))
'''}
Private Sub Class_Initialize()
    
End Sub
