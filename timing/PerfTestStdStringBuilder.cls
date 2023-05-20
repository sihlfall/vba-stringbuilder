VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerfTestStdStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPerfTestCase
Private myNInitial As Long
Private myNToAppend As Long

Public Property Get IPerfTestCase_EnabledByDefault() As Boolean
    IPerfTestCase_EnabledByDefault = False
End Property

Public Property Let IPerfTestCase_NInitial(n As Long)
    myNInitial = n
End Property

Public Property Let IPerfTestCase_NToAppend(n As Long)
    myNToAppend = n
End Property

Public Property Get IPerfTestCase_Description() As String
    IPerfTestCase_Description = "stdStringBuilder"
End Property

Public Sub IPerfTestCase_Run()
    Dim j As Long
    Dim sb As stdStringBuilder
    Set sb = New stdStringBuilder
    sb.JoinStr = vbNullString
    
    sb.Append (String(myNInitial, "a"))
    For j = 1 To myNToAppend
        sb.Append ("b")
    Next
End Sub






