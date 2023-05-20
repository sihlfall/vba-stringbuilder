VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Timing"
   ClientHeight    =   10545
   ClientLeft      =   2340
   ClientTop       =   615
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   7665
   Begin VB.TextBox TextNCharactersToAppend 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   38
      Text            =   "10000"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox TextNInitial 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   36
      Text            =   "3000"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox TextNIterations 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1031
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   34
      Text            =   "1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   3675
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   4185
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   4680
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   26
      Top             =   5175
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   5
      Left            =   240
      TabIndex        =   27
      Top             =   5685
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   6180
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   29
      Top             =   6675
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   30
      Top             =   7185
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   7680
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   10
      Left            =   240
      TabIndex        =   32
      Top             =   8175
      Width           =   375
   End
   Begin VB.CheckBox CheckB 
      Height          =   375
      Index           =   11
      Left            =   240
      TabIndex        =   33
      Top             =   8685
      Width           =   375
   End
   Begin VB.CommandButton ButtonRun 
      Caption         =   "Run ..."
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of characters to append"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   1590
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Initial string length"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   37
      Top             =   990
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of iterations"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   390
      Width           =   1410
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   11
      Left            =   2685
      TabIndex        =   22
      Top             =   8760
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   10
      Left            =   2685
      TabIndex        =   21
      Top             =   8265
      Width           =   1695
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   11
      Left            =   675
      TabIndex        =   20
      Top             =   8760
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   10
      Left            =   675
      TabIndex        =   19
      Top             =   8265
      Width           =   1800
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   9
      Left            =   2685
      TabIndex        =   18
      Top             =   7755
      Width           =   1695
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   9
      Left            =   675
      TabIndex        =   17
      Top             =   7755
      Width           =   1800
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   7
      Left            =   2685
      TabIndex        =   16
      Top             =   6765
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   8
      Left            =   2685
      TabIndex        =   15
      Top             =   7260
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   6
      Left            =   2685
      TabIndex        =   14
      Top             =   6255
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   5
      Left            =   2685
      TabIndex        =   13
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   4
      Left            =   2685
      TabIndex        =   12
      Top             =   5265
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   3
      Left            =   2685
      TabIndex        =   11
      Top             =   4755
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   2
      Left            =   2685
      TabIndex        =   10
      Top             =   4260
      Width           =   1695
   End
   Begin VB.Label LabelResult 
      Alignment       =   1  'Right Justify
      Caption         =   "LabelResult"
      Height          =   255
      Index           =   1
      Left            =   2685
      TabIndex        =   9
      Top             =   3765
      Width           =   1695
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   8
      Left            =   675
      TabIndex        =   8
      Top             =   7260
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   7
      Left            =   675
      TabIndex        =   7
      Top             =   6765
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   6
      Left            =   675
      TabIndex        =   6
      Top             =   6255
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   5
      Left            =   675
      TabIndex        =   5
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   4
      Left            =   675
      TabIndex        =   4
      Top             =   5265
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   3
      Left            =   675
      TabIndex        =   3
      Top             =   4755
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   2
      Left            =   675
      TabIndex        =   1
      Top             =   4260
      Width           =   1800
   End
   Begin VB.Label LabelDescription 
      Caption         =   "LabelDescription"
      Height          =   255
      Index           =   1
      Left            =   675
      TabIndex        =   0
      Top             =   3765
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim globalStringBuilder As StringBuilder

Dim testCases() As IPerfTestCase

Dim isRunning As Boolean


Private Sub ClearResultLabels()
    Dim i As Integer
    For i = LBound(testCases) To UBound(testCases)
        LabelResult(i).Caption = ""
    Next
End Sub

Private Sub ButtonRun_Click()
    Dim i As Integer, j As Integer
    Dim startTime As Double, elapsedTime As Double
    Dim NInitial As Long, NToAppend As Long, NIterations As Long
    
    If isRunning Then Exit Sub
    
    isRunning = True
    ButtonRun.Enabled = False
    DisableCheckboxes
    ClearResultLabels
    
    NInitial = CLng(Me.TextNInitial.Text)
    NIterations = CLng(Me.TextNIterations.Text)
    NToAppend = CLng(Me.TextNCharactersToAppend.Text)
    
    For j = LBound(testCases) To UBound(testCases)
        DoEvents
    
        If testCases(j) Is Nothing Or CheckB(j).Value = 0 Then GoTo Continue
        
        testCases(j).NInitial = NInitial
        testCases(j).NToAppend = NToAppend
        
        startTime = MicroTimer()
        For i = 1 To NIterations
            testCases(j).Run
        Next
        elapsedTime = MicroTimer() - startTime
        LabelResult(j).Caption = Format(elapsedTime * 1000, "0.00") & " ms"
Continue:
    Next
        
    isRunning = False
    ButtonRun.Enabled = True
    EnableCheckboxes
End Sub

Private Sub EnableCheckboxes()
    Dim i As Integer, c As CheckBox
    For i = 1 To 11
        CheckB(i).Enabled = Not testCases(i) Is Nothing
    Next
End Sub

Private Sub DisableCheckboxes()
    Dim c As CheckBox
    For Each c In CheckB
        c.Enabled = False
    Next
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim c As CheckBox

    isRunning = False
    
    ReDim testCases(1 To 11) As IPerfTestCase
    Set testCases(1) = New PerfTestNaive
    Set testCases(2) = New PerfTestDynamic
    Set testCases(3) = New PerfTestStatic
    Set testCases(4) = New PerfTestVolteFaceDragokas
    Set testCases(5) = New PerfTestStdStringBuilder
    Set testCases(6) = New PerfTestStdStringBuilderSsb
    
    For i = LBound(testCases) To UBound(testCases)
        If testCases(i) Is Nothing Then
            LabelDescription(i).Caption = ""
        Else
            LabelDescription(i).Caption = testCases(i).Description
        End If
    Next
    
    ClearResultLabels
    
    EnableCheckboxes
    
    For i = 1 To 11
        If testCases(i) Is Nothing Then
            CheckB(i).Value = 0
        Else
            CheckB(i).Value = -testCases(i).EnabledByDefault
        End If
    Next
    
End Sub

