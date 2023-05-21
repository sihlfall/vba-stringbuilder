Attribute VB_Name = "stdStringBuilderSsbTests"
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
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Append_001()
    ' Test appending several times, staying within minimum buffer size
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    Assert.AreEqual "", sb.Str
    sb.Append ("xyz")
    Assert.AreEqual "xyz", sb.Str
    sb.Append ("abc")
    Assert.AreEqual "xyzabc", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Append_002()
    ' Test appending several times, exceeding minimum buffer size
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(100, "a"))
    Assert.AreEqual String(100, "a"), sb.Str
    sb.Append (String(200, "x"))
    Assert.AreEqual String(100, "a") & String(200, "x"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Append_003()
    ' Test appending several times, exceeding minimum buffer size
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    Assert.AreEqual "", sb.Str
    sb.Append (String(100, "a"))
    Assert.AreEqual String(100, "a"), sb.Str
    sb.Append (String(200, "x"))
    Assert.AreEqual String(100, "a") & String(200, "x"), sb.Str
    sb.Append (String(1000, "y"))
    Assert.AreEqual String(100, "a") & String(200, "x") & String(1000, "y"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Append_004()
    ' Test appending vbNullString to empty buffer
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    Assert.AreEqual "", sb.Str
    sb.Append (vbNullString)
    Assert.AreEqual "", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Append_010()
    ' Test that calling with bracket notation works
    On Error GoTo TestFail
    
    Dim sb As Object
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    Assert.AreEqual "", sb.Str
    sb.[abcdef]
    sb.[ghijkl]
    Assert.AreEqual "abcdefghijkl", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_001()
    ' Test that initially an empty string is returned
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    Assert.AreEqual "", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_002()
    ' Test that the correct string is returned from buffer 0
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(15, "a"))
    Assert.AreEqual "aaaaaaaaaaaaaaa", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_003()
    ' Test that the correct string is returned from buffer 1
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(15, "a"))
    sb.Append (String(8, "b"))
    Assert.AreEqual "aaaaaaaaaaaaaaabbbbbbbb", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_050()
    ' Test: Str is default get property
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb, s As String
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append ("xyzabc")
    s = sb
    Assert.AreEqual "xyzabc", s
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_051()
    ' Test: Str is default let property
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append ("xyzabc")
    sb = "hello world"
    Assert.AreEqual "hello world", sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub StringBuilder_Str_052()
    ' Test: Assigning a short string to Str works if current content is a rather long string
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = String(50, "b")
    Assert.AreEqual String(50, "b"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_053()
    ' Test: After a short string has been assigned to a StringBuilder containing a rather long string,
    '   appending works
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = String(50, "b")
    sb.Append (String(2000, "c"))
    Assert.AreEqual String(50, "b") & String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_054()
    ' Test: Assigning the empty string works
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = ""
    Assert.AreEqual "", sb.Str
    sb.Append (String(2000, "c"))
    Assert.AreEqual String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_055()
    ' Test: Assigning the null string works
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = vbNullString
    Assert.AreEqual "", sb.Str
    sb.Append (String(2000, "c"))
    Assert.AreEqual String(2000, "c"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_056()
    ' Test: Correctly appends single characters after a short string has been assigned to a
    '    StringBuilder containing a long string
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb, i As Integer
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append (String(10000, "a"))
    sb = "abc"
    Assert.AreEqual "abc", sb.Str
    For i = 1 To 2500
        sb.Append ("d")
    Next
    Assert.AreEqual "abc" & String(2500, "d"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_057()
    ' Test: Correctly appends single characters after a long string has been assigned to a
    '    StringBuilder containing a short string
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb, i As Integer
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.Append ("aa")
    Assert.AreEqual "aa", sb.Str
    sb.Str = String(10000, "b")
    Assert.AreEqual String(10000, "b"), sb.Str
    For i = 1 To 25000
        sb.Append ("d")
    Next
    Assert.AreEqual String(10000, "b") & String(25000, "d"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Str_060()
    ' Test: Correctly appends single characters after setting MinimumCapacity = 0
    On Error GoTo TestFail
    
    Dim sb As stdStringBuilderSsb, i As Integer
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = vbNullString
    sb.MinimumCapacity = 0
    Assert.AreEqual 2&, sb.MinimumCapacity
    For i = 1 To 1000
        sb.Append ("d")
    Next
    Assert.AreEqual String(1000, "d"), sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("stdStringBuilderSsb")
Public Sub stdStringBuilderSsb_Demo()
    ' Test: Demo code
    On Error GoTo TestFail
    
    Dim sb As Object
    Set sb = stdStringBuilderSsb.Create()
    sb.JoinStr = "-"
    sb.Str = "Start"
    sb.TrimBehaviour = RTrim
    sb.InjectionVariables.Add "@1", "cool"
    sb.[This is a really cool multi-line    ]
    sb.[string which can even include       ]
    sb.[symbols like " ' # ! / \ without    ]
    sb.[causing compiler errors!!           ]
    sb.[also this has @1 variable injection!]
    Assert.AreEqual "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!", _
        sb.Str
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

