Attribute VB_Name = "StaticStringBuilderTests"
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

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_AppendStr_001()
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, "xyz"
    Assert.AreEqual "xyz", StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, "abc"
    Assert.AreEqual "xyzabc", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_AppendStr_002()
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(100, "a")
    Assert.AreEqual String(100, "a"), StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, String(200, "x")
    Assert.AreEqual String(100, "a") & String(200, "x"), StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_AppendStr_003()
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(100, "a")
    Assert.AreEqual String(100, "a"), StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, String(200, "x")
    Assert.AreEqual String(100, "a") & String(200, "x"), StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, String(1000, "y")
    Assert.AreEqual String(100, "a") & String(200, "x") & String(1000, "y"), StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_AppendStr_004()
    ' Test: Correctly appends null string to null string
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, vbNullString
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, vbNullString
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_AppendStr_005()
    ' Test: Appending empty string works when capacity limit is reached
    On Error GoTo TestFail
        
    Dim sb As StaticStringBuilder.Ty, n As Long
    StaticStringBuilder.AppendStr sb, "a"
    n = sb.Capacity - 1
    StaticStringBuilder.AppendStr sb, String(n, "a")
    Assert.AreEqual sb.Capacity, sb.Length
    StaticStringBuilder.AppendStr sb, vbNullString
    Assert.AreEqual String(n + 1, "a"), StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetStr_001()
    ' Test: Str initially delivers empty string
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetStr_002()
    ' Test: Str delivers correct built string from first buffer
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(15, "a")
    Assert.AreEqual "aaaaaaaaaaaaaaa", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetStr_003()
    ' Test: Str delivers correct built string from second buffer
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(15, "a")
    StaticStringBuilder.AppendStr sb, String(8, "b")
    Assert.AreEqual "aaaaaaaaaaaaaaabbbbbbbb", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_Clear_001()
    ' Test: StaticStringBuilder returns empty string after calling Clear
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(15, "a")
    Assert.AreEqual String(15, "a"), StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.Clear sb
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_Clear_002()
    ' Test: Appending after calling Clear works
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(15, "a")
    StaticStringBuilder.Clear sb
    StaticStringBuilder.AppendStr sb, String(50, "b")
    Assert.AreEqual String(50, "b"), StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_Clear_003()
    ' Test: Clear can be called on empty StaticStringBuilder
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.Clear sb
    Assert.AreEqual "", StaticStringBuilder.GetStr(sb)
    StaticStringBuilder.AppendStr sb, String(50, "b")
    Assert.AreEqual String(50, "b"), StaticStringBuilder.GetStr(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetLength_001()
    ' Test: GetLength delivers 0 for initially empty StaticStringBuilder
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    Assert.AreEqual 0&, StaticStringBuilder.GetLength(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetLength_002()
    ' Test: GetLength delivers correct length after appending
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, String(50, "a")
    Assert.AreEqual 50&, StaticStringBuilder.GetLength(sb)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_001()
    ' Test: GetSubstr delivers empty string for initially empty StaticStringBuilder
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    Assert.AreEqual "", StaticStringBuilder.GetSubstr(sb, 5, 2)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_002()
    ' Test: GetSubstr returns empty string when start = string length + 1
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "", StaticStringBuilder.GetSubstr(sb, 5, 2)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_003()
    ' Test: GetSubstr returns last character when start = string length
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "d", StaticStringBuilder.GetSubstr(sb, 4, 2)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_004()
    ' Test: GetSubstr returns entire string when start = 1 and length > length of string
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "abcd", StaticStringBuilder.GetSubstr(sb, 1, 10)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_005()
    ' Test: GetSubstr returns entire string when start = 1 and length = length of string
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "abcd", StaticStringBuilder.GetSubstr(sb, 1, 4)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_006()
    ' Test: GetSubstr returns entire string without last character when start = 1 and length = length of string - 1
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "abc", StaticStringBuilder.GetSubstr(sb, 1, 3)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_GetSubstr_007()
    ' Test: GetSubstr returns empty string when length = 0
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "abcd"
    Assert.AreEqual "", StaticStringBuilder.GetSubstr(sb, 1, 0)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_SetMinimumCapacity_001()
    ' Test: StaticStringBuilder initially uses correct capacity after SetMinimumCapacity
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.SetMinimumCapacity sb, 100
    StaticStringBuilder.AppendStr sb, "a"
    Assert.AreEqual 100&, sb.Capacity
    Assert.AreEqual 100&, Len(sb.Buffer(sb.Active))
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_SetMinimumCapacity_002()
    ' Test: StaticStringBuilder correctly increases capacity after SetMinimumCapacity
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.SetMinimumCapacity sb, 100
    StaticStringBuilder.AppendStr sb, "a"
    StaticStringBuilder.AppendStr sb, String(100, "a")
    Assert.AreEqual 150&, sb.Capacity
    Assert.AreEqual 150&, Len(sb.Buffer(sb.Active))
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_SetMinimumCapacity_003()
    ' Test: StaticStringBuilder uses new minimum capacity on increase after SetMinimumCapacity
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "a"
    StaticStringBuilder.SetMinimumCapacity sb, 100
    StaticStringBuilder.AppendStr sb, String(50, "a")
    Assert.AreEqual 100&, sb.Capacity
    Assert.AreEqual 100&, Len(sb.Buffer(sb.Active))
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("StaticStringBuilder")
Public Sub StaticStringBuilder_Demo()
    ' Test: Demo code from README file
    On Error GoTo TestFail
    
    Dim sb As StaticStringBuilder.Ty
    StaticStringBuilder.AppendStr sb, "First"
    StaticStringBuilder.AppendStr sb, "Second"
    StaticStringBuilder.AppendStr sb, "Third"
    
    Dim s As String
    s = StaticStringBuilder.GetStr(sb)
    Assert.AreEqual "FirstSecondThird", s
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

