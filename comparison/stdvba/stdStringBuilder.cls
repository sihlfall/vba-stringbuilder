VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'StringBuilder
'
'StringBuilder is specifically designed to build large strings which are mainly made from boiler plate code.
'Instead of code like this:
'
'    Dim str as string
'    str = str & "<div className=""MarkdownEditor"">" & vbCrLf
'    str = str & "  <h3>Input</h3>" & vbCrLf
'    str = str & "  <label htmlFor=""markdown-content"">" & vbCrLf
'    str = str & "    Enter some markdown" & vbCrLf
'    str = str & "  </label>" & vbCrLf
'    str = str & "  <textarea" & vbCrLf
'    str = str & "    id=""markdown-content""" & vbCrLf
'    str = str & "    onChange=""" & handleChange & """" & vbCrLf & _ " & vbCrLf
'    str = str & "    defaultValue=""" & state.value & """" & vbCrLf
'    str = str & "  />" & vbCrLf
'    str = str & "  <h3>Output</h3>" & vbCrLf
'    str = str & "  <div" & vbCrLf
'    str = str & "    className=""content""" & vbCrLf
'    str = str & "    dangerouslySetInnerHTML=" & getRawMarkup() & vbCrLf
'    str = str & "  />" & vbCrLf
'    str = str & "</div>"
'    renderHTML(str)
'
'You can now use code like this:
'    'IMPORTANT!!! Only Object (aka "IDispatch") can use square bracket syntax!
'    Dim sb as Object
'    set sb = StringBuilder.Create()
'    sb.TrimBehaviour = RTrim
'
'    'Inject variables into string
'    sb.InjectionVariables.add "{this.handleChange}", handleChange
'    sb.InjectionVariables.add "{this.state.value}", state.value
'    sb.InjectionVariables.add "{this.getRawMarkup()}", getRawMarkup()
'
'    'Build string
'    sb.[<div className="MarkdownEditor">                 ]
'    sb.[  <h3>Input</h3>                                 ]
'    sb.[  <label htmlFor="markdown-content">             ]
'    sb.[    Enter some markdown                          ]
'    sb.[  </label>                                       ]
'    sb.[  <textarea                                      ]
'    sb.[    id="markdown-content"                        ]
'    sb.[    onChange="{this.handleChange}"               ]
'    sb.[    defaultValue="{this.state.value}"            ]
'    sb.[  />                                             ]
'    sb.[  <h3>Output</h3>                                ]
'    sb.[  <div                                           ]
'    sb.[    className="content"                          ]
'    sb.[    dangerouslySetInnerHTML={this.getRawMarkup()}]
'    sb.[  />                                             ]
'    sb.[</div>                                           ]
'    renderHTML(sb)
'
'Which significantly improves maintainability of the code.

Private Tainted As Boolean
Private ProcessedString As String
Public RawString As String
Public JoinStr As String
Public TrimBehaviour As TrimBehaviourEnum
Public Enum TrimBehaviourEnum
  NoTrim
  LTrim
  RTrim
  Trim
End Enum
Public InjectionVariables As Object
Public Property Get Str() As String
Attribute Str.VB_UserMemId = 0
  If Tainted Then Call RefreshFromRaw
  Str = ProcessedString
End Property
Public Property Let Str(s As String)
  RawString = s
End Property

'Appends the string parsed to the main string Str
Public Function Append(s As String) As Variant
Attribute Append.VB_UserMemId = -5
  'Trim based on TrimBehaviour
  Select Case TrimBehaviour
    Case TrimBehaviourEnum.LTrim
      s = VBA.LTrim(s)
    Case TrimBehaviourEnum.RTrim
      s = VBA.RTrim(s)
    Case TrimBehaviourEnum.Trim
      s = VBA.Trim(s)
  End Select
  
  'Build String
  RawString = RawString & JoinStr & s
  
  'Set tainted
  Tainted = True
End Function

'Creates an instance of stdStringBuilder
Public Function Create() As stdStringBuilder
  Set Create = New stdStringBuilder
End Function

Public Function Test()
  Dim sb As Object
  Set sb = stdStringBuilder.Create()
  sb.JoinStr = "-"
  sb.Str = "Start"
  sb.TrimBehaviour = RTrim
  sb.InjectionVariables.Add "@1", "cool"
  sb.[This is a really cool multi-line    ]
  sb.[string which can even include       ]
  sb.[symbols like " ' # ! / \ without    ]
  sb.[causing compiler errors!!           ]
  sb.[also this has @1 variable injection!]
  Test = sb.Str = "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!"
End Function

'==============================
' PRIVATE METHODS
'==============================

Private Sub RefreshFromRaw()
  ProcessedString = RawString
  
  'Replace keys
  Dim key
  For Each key In InjectionVariables.Keys()
    ProcessedString = VBA.Replace(ProcessedString, key, InjectionVariables(key))
  Next
  
  'No longer tainted
  Tainted = False
End Sub
Private Sub Class_Initialize()
  Str = ""
  JoinStr = vbCrLf
  Set InjectionVariables = CreateObject("Scripting.Dictionary")
End Sub
