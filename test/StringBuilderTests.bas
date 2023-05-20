Attribute VB_Name = "StringBuilderTests"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_001()
    On Error GoTo TestFail
    
    With New StringBuilder
        Assert.AreEqual "", .Str
        .AppendStr "xyz"
        Assert.AreEqual "xyz", .Str
        .AppendStr "abc"
        Assert.AreEqual "xyzabc", .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_002()
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr String(100, "a")
        Assert.AreEqual String(100, "a"), .Str
        .AppendStr String(200, "x")
        Assert.AreEqual String(100, "a") & String(200, "x"), .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_003()
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr String(100, "a")
        Assert.AreEqual String(100, "a"), .Str
        .AppendStr String(200, "x")
        Assert.AreEqual String(100, "a") & String(200, "x"), .Str
        .AppendStr String(1000, "y")
        Assert.AreEqual String(100, "a") & String(200, "x") & String(1000, "y"), .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_004()
    ' Test: Correctly appends null string to null string
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr vbNullString
    Assert.AreEqual "", sb.Str
    sb.AppendStr vbNullString
    Assert.AreEqual "", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_005()
    ' Test: Appending empty string works when capacity limit is reached
    On Error GoTo TestFail
        
    Dim sb As StringBuilder, n As Long
    Set sb = New StringBuilder
    With sb
        .AppendStr "a"
        n = .Capacity - 1
        .AppendStr String(n, "a")
        Assert.AreEqual .Capacity, .Length
        .AppendStr vbNullString
        Assert.AreEqual String(n + 1, "a"), sb.Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_AppendStr_010()
    ' Bracket notation works
    On Error GoTo TestFail
    
    Dim sb As Object
    Set sb = New StringBuilder
    sb.[abc]
    sb.[def]
    Assert.AreEqual "abcdef", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_001()
    ' Test: Str initially delivers empty string
    On Error GoTo TestFail
    
    With New StringBuilder
        Assert.AreEqual "", .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_002()
    ' Test: Str delivers correct built string from first buffer
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr String(15, "a")
        Assert.AreEqual "aaaaaaaaaaaaaaa", .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_003()
    ' Test: Str delivers correct built string from second buffer
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr String(15, "a")
        .AppendStr String(8, "b")
        Assert.AreEqual "aaaaaaaaaaaaaaabbbbbbbb", .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_050()
    ' Test: Str is default get property
    On Error GoTo TestFail
    
    Dim sb As StringBuilder, s As String
    Set sb = New StringBuilder
    sb.AppendStr "xyzabc"
    s = sb
    Assert.AreEqual "xyzabc", s
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_051()
    ' Test: Str is default let property
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr "xyzabc"
    sb = "hello world"
    Assert.AreEqual "hello world", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_052()
    ' Test: Assigning a short string to Str works if current content is a rather long string
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr String(10000, "a")
    sb = String(50, "b")
    Assert.AreEqual String(50, "b"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_053()
    ' Test: After a short string has been assigned to a StringBuilder containing a rather long string,
    '   appending works
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr String(10000, "a")
    sb = String(50, "b")
    sb.AppendStr String(2000, "c")
    Assert.AreEqual String(50, "b") & String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_054()
    ' Test: Assigning the empty string works
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr String(10000, "a")
    sb = ""
    Assert.AreEqual "", sb.Str
    sb.AppendStr String(2000, "c")
    Assert.AreEqual String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_055()
    ' Test: Assigning the null string works
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr String(10000, "a")
    sb = vbNullString
    Assert.AreEqual "", sb.Str
    sb.AppendStr String(2000, "c")
    Assert.AreEqual String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Str_056()
    ' Test: Correctly appends single characters after a short string has been assigned to a
    '    StringBuilder containing a long string
    On Error GoTo TestFail
    
    Dim sb As StringBuilder, i As Integer
    Set sb = New StringBuilder
    sb.AppendStr String(10000, "a")
    sb = "abc"
    Assert.AreEqual "abc", sb.Str
    For i = 1 To 2500
        sb.AppendStr ("d")
    Next
    Assert.AreEqual "abc" & String(2500, "d"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Clear_001()
    ' Test: StringBuilder returns empty string after calling Clear
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr (String(15, "a"))
        Assert.AreEqual String(15, "a"), .Str
        .Clear
        Assert.AreEqual "", .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Clear_002()
    ' Test: Appending after calling Clear works
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr (String(15, "a"))
        .Clear
        .AppendStr (String(50, "b"))
        Assert.AreEqual String(50, "b"), .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Clear_003()
    ' Test: Clear can be called on empty StringBuilder
    On Error GoTo TestFail
    
    With New StringBuilder
        .Clear
        Assert.AreEqual "", .Str
        .AppendStr String(50, "b")
        Assert.AreEqual String(50, "b"), .Str
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Length_001()
    ' Test: Length delivers 0 for initially empty StringBuilder
    On Error GoTo TestFail
    
    With New StringBuilder
        Assert.AreEqual 0&, .Length
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Length_002()
    ' Test: Length delivers correct length after appending
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr (String(50, "a"))
        Assert.AreEqual 50&, .Length
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_001()
    ' Test: Substr delivers empty string for initially empty StringBuilder
    On Error GoTo TestFail
    
    With New StringBuilder
        Assert.AreEqual "", .Substr(5, 2)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_002()
    ' Test: Substr returns empty string when start = string length + 1
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "", .Substr(5, 2)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_003()
    ' Test: Substr returns last character when start = string length
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "d", .Substr(4, 2)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_004()
    ' Test: Substr returns entire string when start = 1 and length > length of string
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "abcd", .Substr(1, 10)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_005()
    ' Test: Substr returns entire string when start = 1 and length = length of string
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "abcd", .Substr(1, 4)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_006()
    ' Test: Substr returns entire string without last character when start = 1 and length = length of string - 1
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "abc", .Substr(1, 3)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Substr_007()
    ' Test: Substr returns empty string when length = 0
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "abcd"
        Assert.AreEqual "", .Substr(1, 0)
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_SetMinimumCapacity_001()
    ' Test: StringBuilder initially uses correct capacity after SetMinimumCapacity
    On Error GoTo TestFail
    
    With New StringBuilder
        .MinimumCapacity = 100
        .AppendStr "a"
        Assert.AreEqual 100&, .Capacity
        Assert.AreEqual 100&, .MinimumCapacity
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_SetMinimumCapacity_002()
    ' Test: StringBuilder correctly increases capacity after SetMinimumCapacity
    On Error GoTo TestFail
    
    With New StringBuilder
        .MinimumCapacity = 100
        .AppendStr "a"
        .AppendStr String(100, "a")
        Assert.AreEqual 150&, .Capacity
        Assert.AreEqual 100&, .MinimumCapacity
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_SetMinimumCapacity_003()
    ' Test: StringBuilder uses new minimum capacity on increase after SetMinimumCapacity
    On Error GoTo TestFail
    
    With New StringBuilder
        .AppendStr "a"
        .MinimumCapacity = 100
        .AppendStr String(50, "a")
        Assert.AreEqual 100&, .Capacity
        Assert.AreEqual 100&, .MinimumCapacity
    End With
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Demo_001()
    ' Test: Demo code from README file
    On Error GoTo TestFail
    
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    sb.AppendStr "First"
    sb.AppendStr "Second"
    sb.AppendStr "Third"
    Dim s As String
    s = sb.Str
    Assert.AreEqual "FirstSecondThird", s
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StringBuilder")
Public Sub StringBuilder_Demo_002()
    ' Test: Demo code from README file
    Dim sb As Object, s As String
    Set sb = New StringBuilder
    sb.[First]
    sb.[Second]
    sb.[Third]
    s = sb
    Assert.AreEqual "FirstSecondThird", s
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

