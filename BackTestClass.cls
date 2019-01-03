VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public gobj As SystemClass
Public FromExpect As String
Public ToExpect As String
Public CheckType As Integer
Public ChipCount As Integer
Public MinCheckTimes As Integer
Public StepLen As Integer
Public ReadDataCount As Long
Const BefDataCount = 180 '前留数据180


Public Function ExecTest(ByRef RepeatCodeDic As Dictionary)
    Dim i As Long
    Dim data() As ExpectData
    Dim ed As New ExpectData
    Dim Chances() As ChanceClass
    Dim HoldChance() As Dictionary
    Dim OpenCodeDic As New Dictionary
    Dim SameCodeDic As New Dictionary
    ed.TableToArray ed.GetHistoryData(Me.FromExpect, Me.ReadDataCount, "asc"), data
    Dim c As Long
    i = 1
    c = 0
    Dim preExpect As String
    While i <= UBound(data)
        Dim currExpect As String, currOpenCode As String
        currExpect = data(i).Expect
        currOpenCode = data(i).OpenCode
        If Not OpenCodeDic.Exists(currOpenCode) Then
            OpenCodeDic.Add currOpenCode, currExpect
        Else
            Dim sameExpects() As String
            If SameCodeDic.Exists(currOpenCode) Then
                sameExpects = SameCodeDic(currOpenCode)
                Dim samecnt As Integer
                samecnt = UBound(sameExpects)
                ReDim Preserve sameExpects(samecnt + 1)
                sameExpects(samecnt + 1) = currExpect
                SameCodeDic(currOpenCode) = sameExpects
            Else
                ReDim sameExpects(2)
                sameExpects(1) = OpenCodeDic(currOpenCode)
                sameExpects(2) = currExpect
                SameCodeDic.Add currOpenCode, sameExpects
            End If
        End If
        i = i + 1
        c = c + 1
        If i > UBound(data) Then
            Dim tmpdata() As ExpectData
            ed.TableToArray ed.GetHistoryData(currExpect, Me.ReadDataCount, "asc"), tmpdata
            If UBound(tmpdata) = 0 Then '如果找不到数据，退出
                GoTo breakWhile
            End If
            Dim befdata() As ExpectData
            befdata = data
            ReDim data(0)
            MergeData befdata, tmpdata, BefDataCount, data
            i = BefDataCount + 1
        End If
    Wend
breakWhile:
Set RepeatCodeDic = SameCodeDic
'MsgBox "共回测了" & CStr(c) & "条历史数据！"
End Function

Sub MergeData(beforedata() As ExpectData, lastdata() As ExpectData, befcnt As Long, ByRef outdata() As ExpectData)
    Dim i As Long
    Dim shift As Long
    shift = UBound(beforedata) - befcnt
    ReDim outdata(UBound(lastdata) + befcnt)
    For i = 1 To befcnt
        Set outdata(i) = New ExpectData
        outdata(i).Expect = beforedata(shift + i).Expect
        outdata(i).OpenCode = beforedata(shift + i).OpenCode
        outdata(i).OpenTime = beforedata(shift + i).OpenTime
    Next
    For i = 1 To UBound(lastdata)
        Set outdata(befcnt + i) = New ExpectData
        outdata(befcnt + i).Expect = lastdata(i).Expect
        outdata(befcnt + i).OpenCode = lastdata(i).OpenCode
        outdata(befcnt + i).OpenTime = lastdata(i).OpenTime
    Next
End Sub


Public Function TestSimpleType()
    
End Function
