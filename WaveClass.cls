VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaveClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim AllData() As ExpectData
Public CheckDays As Integer '50次
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Function getData(source() As ExpectData, Frm As Long, lng As Long) As ExpectData()
    Dim ret() As ExpectData
    Dim reallng As Long
    reallng = lng
    If UBound(source) < Frm + lng Then
        reallng = UBound(source) - Frm + 1
    End If
    ReDim ret(reallng)
    Dim i As Long
    For i = Frm To Frm + reallng - 1
        Set ret(i - Frm + 1) = source(i)
    Next
    getData = ret
End Function
Public Sub init(data() As ExpectData)
    AllData = data
End Sub

Function EMA(data() As Single, n As Integer, Optional index As Integer)
    If index = UBound(data) Then
        EMA = data(index)
    Else
        EMA = data(index) + 0.5 * EMA(data(), n, index + 1)
    End If
End Function

Function CalcChartData(data() As ExpectData) As Dictionary
    Dim ret As New Dictionary
    Dim i As Integer
    Dim AllSerialSum As Long
    Dim cm As New ClassModel
    For i = 1 To 10
        Dim j As Integer
        Dim ExistCars As Dictionary '车的信息，主要是连续出现次数
        Set ExistCars = New Dictionary
        Dim RepeatCnt As Integer
        RepeatCnt = 0
        Dim lastCar As String
        Dim NoCarsArr() As Single '前推各期缺失车数组
        ReDim NoCarsArr(UBound(data))
        For j = 1 To UBound(data)
            Dim posDir As Dictionary
            Set posDir = New Dictionary
            
            Dim car As String
            Dim cars As String
            cars = data(j).OpenCode
            car = Split(cars, ",")(i - 1)
            Dim CarRepeatCnt() As Integer
            ReDim CarRepeatCnt(0)
            If Not ExistCars.Exists(car) Then
                ExistCars.Add car, CarRepeatCnt
            End If
            CarRepeatCnt = ExistCars(car)
            If j = 1 Then
                
            Else
                Dim currCnt As Integer
                If car = lastCar Then
                    RepeatCnt = RepeatCnt + 1
                Else
                    CarRepeatCnt = ExistCars(lastCar)
                    currCnt = UBound(CarRepeatCnt)
                    ReDim Preserve CarRepeatCnt(currCnt + 1)
                    CarRepeatCnt(currCnt + 1) = RepeatCnt + 1
                    ExistCars(lastCar) = CarRepeatCnt
                    RepeatCnt = 0
                End If
            End If
            
            If j = UBound(data) Then
                CarRepeatCnt = ExistCars(car)
                currCnt = UBound(CarRepeatCnt)
                ReDim Preserve CarRepeatCnt(currCnt + 1)
                CarRepeatCnt(currCnt + 1) = RepeatCnt + 1
                ExistCars(car) = CarRepeatCnt
            End If
            lastCar = car
        Next
        Dim key As String
        Dim ExistCarWeight As Dictionary
        Set ExistCarWeight = New Dictionary
        Dim AllSum As Integer
        AllSum = 0
        For Each kk In ExistCars.Keys
            key = kk
            Dim sum As Integer
            sum = 0
            Dim CntArr() As Integer
            CntArr = ExistCars(key)
            Dim c As Integer
            For c = 1 To UBound(CntArr)
                sum = sum + CntArr(c) ^ 2
            Next
            If ExistCarWeight.Exists(key) = False Then
                ExistCarWeight.Add key, sum
            End If
            AllSum = AllSum + sum
        Next
        Dim CurrArea As Single
        CurrArea = getNoCarsArea(data, i)
        ret.Add i, AllSum / CurrArea
        AllSerialSum = AllSerialSum + AllSum
    Next
    ret.Add 0, AllSerialSum / 10
    Set CalcChartData = ret
End Function

Function getNoCarsArea(data() As ExpectData, serno As Integer) As Single
    Dim cntSum As Integer '所有未出现车的次数
    Dim maxCnt As Integer '最大未出现次数
    Dim i As Integer
    For i = 1 To 10
       Dim lastCar As String
       Dim j As Integer
       Dim jCnt As Integer
       jCnt = 0
       For j = 1 To UBound(data)
            Dim car As String
            car = Split(data(j).OpenCode, ",")(serno - 1)
            If i = CInt(car) Then
                jCnt = j - 1
                GoTo End_for_getNoCarsCount_j
            End If
       Next
End_for_getNoCarsCount_j:
       cntSum = cntSum + jCnt
       If jCnt > maxCnt Then maxCnt = jCnt
    Next
    getNoCarsArea = (10 * maxCnt - cntSum) / 10 '最大深度*10-未出现的车/10
End Function
