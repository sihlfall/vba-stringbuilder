VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringBuilder
' v0.1.2 by Sihlfall
' MIT license
'
' A (relatively) performant, portable StringBuilder class for VBA / VB6.
'
' Usage:
' Copy or import this class module into your project.
'
' Dim sb As StringBuilder
' Set sb = New StringBuilder
' sb.AppendStr "First"
' sb.AppendStr "Second"
' sb.AppendStr "Third"
'
' Dim s As String
' s = sb.Str
' ' Now s equals "FirstSecondThird"
'
Option Explicit


' Must be at least 2.
Private Const DEFAULT_MINIMUM_CAPACITY As Long = 16

Private mActive As Integer            ' index of the currently active buffer (0 or 1)
Private mBuffer(0 To 1) As String     ' mBuffer(.active) is the currently active buffer
Private mCapacity As Long             ' current allocated capacity in characters
Private mLength As Long               ' current length of the string, in characters
Private mMinimumCapacity As Long      ' minimum capacity set (0 or >= 2)

''' Public properties

Public Property Get Capacity() As Long
    Capacity = mCapacity
End Property

Public Property Get Length() As Long
    Length = mLength
End Property

Public Property Get MinimumCapacity() As Long
    MinimumCapacity = mMinimumCapacity
End Property

Public Property Let MinimumCapacity(ByVal desiredMinimumCapacity As Long)
    If desiredMinimumCapacity >= 2 Then mMinimumCapacity = desiredMinimumCapacity Else mMinimumCapacity = 2
End Property

' Str is configured to be the default property, so you should be able to write the following:
'   dim s as string, sb as StringBuilder
'   ...
'   s = sb  ' equivalent to s = sb.Str
Public Property Get Str() As String
Attribute Str.VB_UserMemId = 0
    Str = Left$(mBuffer(mActive), mLength)
End Property

Public Property Let Str(ByRef s As String)
    Dim sLength As Long
    sLength = Len(s)
    ' We accept only 100% space overhead.
    If mCapacity < sLength Or (2 * sLength < mCapacity And mCapacity > mMinimumCapacity) Then ClearReserve sLength
    If sLength > 0 Then Mid$(mBuffer(mActive), 1, sLength) = s
    mLength = sLength
End Property

Public Property Get Substr(ByVal start As Long, ByVal Length As Long) As String
    Dim n As Long
    n = mLength - start + 1
    If n <= 0 Then
        Substr = vbNullString
        Exit Property
    End If
    If Length <= n Then n = Length
    Substr = Mid$(mBuffer(mActive), start, n)
End Property

''' Public subroutines and functions

' Append s to the string being built.
' s is taken by reference for performance reasons only. s will remain unchanged.
Public Sub AppendStr(ByRef s As String)
Attribute AppendStr.VB_UserMemId = -5
    Dim Length As Long, nRequired As Long
    Length = Len(s)
    If Length = 0 Then Exit Sub
    nRequired = mLength + Length
    If nRequired > mCapacity Then SwitchToLargerBuffer nRequired
    Mid$(mBuffer(mActive), mLength + 1, Length) = s
    mLength = nRequired
End Sub

Public Sub Clear()
    mActive = 0
    mBuffer(0) = vbNullString
    mBuffer(1) = vbNullString
    mCapacity = 0
    mLength = 0
End Sub

''' Private subroutines and functions

Private Sub ClearReserve(ByVal nRequired As Long)
    ' Clear everything
    ' Allocate a buffer that is able to hold nRequired characters.
    ' The new buffer size is calculated by repeatedly growing the minimum capacity by 50%.
    mActive = 0
    mLength = 0
    mBuffer(1) = vbNullString
    If nRequired > 0 Then
        mCapacity = mMinimumCapacity
        Do
            If mCapacity >= nRequired Then Exit Do
            mCapacity = mCapacity + mCapacity \ 2
        Loop
        mBuffer(0) = String(mCapacity, 0)
    Else
        mCapacity = 0
        mBuffer(0) = vbNullString
    End If
End Sub

Private Sub SwitchToLargerBuffer(ByVal nRequired As Long)
    ' Allocate buffer that is able to hold nRequired characters.
    ' The new buffer size is calculated by repeatedly growing the current size by 50%.
    ' Copy string over to the new buffer.
    ' Deallocate the old buffer.
    If mCapacity < mMinimumCapacity Then mCapacity = mMinimumCapacity
    Do
        If mCapacity >= nRequired Then Exit Do
        mCapacity = mCapacity + mCapacity \ 2
    Loop
    mBuffer(1 - mActive) = String(mCapacity, 0)
    Mid$(mBuffer(1 - mActive), 1, mLength) = mBuffer(mActive)
    mBuffer(mActive) = vbNullString
    mActive = 1 - mActive
End Sub

Private Sub Class_Initialize()
    mMinimumCapacity = DEFAULT_MINIMUM_CAPACITY
End Sub
