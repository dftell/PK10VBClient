VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CurrDate As Date
Public InitCash As Long
Public CurrCash As Long
Public Gained As Long
Public TotalAsset As Long
Public OneChanceMaxLossRateToInit As Single
Public OneChanceMaxLossRateToTotal As Single
Public MaxLossAbsToTotal As Single
Public MaxGainAbsToInit As Single
'判断持有项目是否可以继续持有
Public Function AllowContinueHold(cc As ChanceClass) As Boolean
    If cc.ChipCount = 1 Then
        If (cc.Cost / InitCash) > OneChanceMaxLossRateToInit Then
            AllowContinueHold = False
            Exit Function
        End If
        AllowContinueHold = AbsLossCheck()
        Exit Function
    End If
    If (cc.Cost / InitCash) > OneChanceMaxLossRateToTotal / 100 Then
        AllowContinueHold = False
        Exit Function
    End If
    AllowContinueHold = AbsLossCheck()
End Function


'判断某机会是否可以下注,掷毂
Public Function AllowBuyByRnd(CurrTimes As Integer, MinTimes As Integer, fromi As Integer, toi As Integer) As Boolean
    AllowBuyByRnd = True
    Exit Function
    
    Math.Randomize
    Dim myTimes As Integer
    myTimes = Int((toi * Math.Rnd) + fromi)
    If CurrTimes < 60 Then
        If CurrTimes > MinTimes + myTimes Then '第一轮
            AllowBuyByRnd = True
            Exit Function
        End If
    Else
        If CurrTimes >= 60 + 5 * myTim / 3 Then '第二轮，专门针对单码
            AllowBuyByRnd = True
            Exit Function
        End If
    End If
    AllowBuyByRnd = False
End Function





'判断全局是否可以新下注
Public Function AllowNewBuyForGlobel(chips As Integer)
''    hrs = Hour(Now())
''    If chips = 1 And (hrs < 9 Or hrs >= 22) Then
''        AllowNewBuyForGlobel = False
''        Exit Function
''    End If
''    If chips >= 3 And (hrs < 9 Or hrs >= 23) Then
''        AllowNewBuyForGlobel = False
''        Exit Function
''    End If
    AllowNewBuyForGlobel = AbsLossCheck()
End Function

Function AbsLossCheck()
    On Error Resume Next

''    If Hour(Now()) = 23 And Minute(Now()) > 58 Then
''        Exit Function
''    End If
''    If (TotalAsset - InitCash) < 0 And Math.Abs((TotalAsset - InitCash)) / InitCash > (MaxLossAbsToTotal / 100) Then
''        AbsLossCheck = False
''        Exit Function
''    End If
    AbsLossCheck = True
End Function

