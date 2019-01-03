VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub ToArray(obj As Object, NodePath As String, ByRef ret() As Object)
    On Error Resume Next
    Dim i As Long
    Dim Length As Long
    Dim str As String
    Dim strPath As String
    Length = CallByName(obj, NodePath & ".Count", VbGet)
    Dim X As New ScriptControl
    X.Language = "JScript"
    X.AddObject "obj", obj
    'Set GetJsonVal = x.Run("j", str)
    i = 0
    While True
        Dim item As Object
        Set item = Nothing
        strPath = "obj." & NodePath & Replace("[i]", "i", i)
        Set item = X.Eval(strPath)
        If item Is Nothing Then
            Exit Sub
        End If
        i = i + 1
        ReDim Preserve ret(i)
        Set ret(i) = item
        'Set item = Nothing
    Wend
End Sub

Public Function GetJsonVal(str As String, root As String) As Object
    On Error Resume Next
    Dim X As New ScriptControl
    'aa = "{ ""people"": [{ ""firstName"": ""Brett"", ""lastName"":""McLaughlin"", ""email"": ""brettnewInstance.com"" },{ ""firstName"": ""Jason"", ""lastName"":""Hunter"", ""email"": ""jasonservlets.com"" }, { ""firstName"": ""Elliotte"", ""lastName"":""Harold"", ""email"": ""elharomacfaq.com"" }]}"
    'Set x = CreateObject("ScriptControl")
    X.Language = "JScript"
    X.AddObject "jsonobj", obj
    s = "function j(s) { obj =eval('('+s+')');return obj;}"
    X.AddCode s
    Set obj = X.Run("j", str)
    If root = "" Then
        Set GetJsonVal = obj
        Exit Function
    End If
    Set X = New ScriptControl
    X.Language = "JScript"
    X.AddObject "jsobj", obj
    s = "jsobj." & root & ""
    Set GetJsonVal = X.Eval(s)
End Function
Public Function RemoveSpeaker(strText)
    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    Dim msg As String
    regex.Pattern = "@[a-f0-9]*:"
    msg = regex.Replace(strText, "")
    RemoveSpeaker = msg
End Function
Public Function RemoveHTML(strText)
    Dim regex
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    Htmlstring = Replace(Replace(Replace(Replace(Replace(strText, "&lt;", "<"), "&gt;", ">"), "&amp;", "&"), "&quot;", """"), "&copy;", "@")
     '//É¾³ý½Å±¾
  'Htmlstring = regex.Replace(Htmlstring, "<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase)
''' &lt; &gt;&amp;&quot;&copy;
''' <£¬>£¬&£¬"£¬?;µÄ×ªÒå×Ö·û
  regex.Pattern = "<script[^>]*?>.*?</script>"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "<[^>]*>"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "([\r\n])[\s]+"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "-->"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "<!--.*"
  Htmlstring = regex.Replace(Htmlstring, "")
  
  regex.Pattern = "&(quot|#34)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(amp|#38)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(lt|#60)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(gt|#62)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(nbsp|#160)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(nbsp|#161)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(nbsp|#162)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(nbsp|#163)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&(nbsp|#169)"
  Htmlstring = regex.Replace(Htmlstring, "")
  regex.Pattern = "&#(\d+)"
  Htmlstring = regex.Replace(Htmlstring, "")
''''''  '//É¾³ýHTML
''''''  Htmlstring = regex.Replace(Htmlstring, "<(.[^>]*)>", "", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "([\r\n])[\s]+", "", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "-->", "", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "<!--.*", "", RegexOptions.IgnoreCase)
''''''
''''''  Htmlstring = regex.Replace(Htmlstring, "&(quot|#34)", """", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(amp|#38)", "&", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(lt|#60)", "<", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(gt|#62)", ">", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(nbsp|#160)", "   ", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(iexcl|#161)", "\xa1", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(cent|#162)", "\xa2", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(pound|#163)", "\xa3", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&(copy|#169)", "\xa9", RegexOptions.IgnoreCase)
''''''  Htmlstring = regex.Replace(Htmlstring, "&#(\d+)", "", RegexOptions.IgnoreCase)
    RemoveHTML = Htmlstring
End Function






