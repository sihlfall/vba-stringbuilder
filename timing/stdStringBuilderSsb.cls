VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stdStringBuilderSsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'StringBuilderSsb
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
Private ssb As StaticStringBuilder.Ty

' Public RawString As String
Public JoinStr As String
Public TrimBehaviour As SsbTrimBehaviourEnum
Public Enum SsbTrimBehaviourEnum
  NoTrim
  LTrim
  RTrim
  Trim
End Enum
Private myInjectionVariables As Object
Public Property Get InjectionVariables() As Object
    If myInjectionVariables Is Nothing Then Set myInjectionVariables = CreateObject("Scripting.Dictionary")
    Set InjectionVariables = myInjectionVariables
End Property

Public Property Get str() As String
  If Tainted Then Call RefreshFromRaw
  str = ProcessedString
End Property
Public Property Let str(s As String)
  StaticStringBuilder.Clear ssb
  StaticStringBuilder.AppendString ssb, s
End Property

'Appends the string parsed to the main string Str
Public Function Append(s As String) As Variant
  'Trim based on TrimBehaviour
  Select Case TrimBehaviour
    Case SsbTrimBehaviourEnum.LTrim
      s = VBA.LTrim(s)
    Case SsbTrimBehaviourEnum.RTrim
      s = VBA.RTrim(s)
    Case SsbTrimBehaviourEnum.Trim
      s = VBA.Trim(s)
  End Select
  
  'Build String
  StaticStringBuilder.AppendString ssb, JoinStr
  StaticStringBuilder.AppendString ssb, s
  
  'Set tainted
  Tainted = True
End Function

'Creates an instance of stdStringBuilderSsb
Public Function Create() As stdStringBuilderSsb
  Set Create = New stdStringBuilderSsb
End Function

Public Function Test()
  Dim sb As Object
  Set sb = stdStringBuilderSsb.Create()
  sb.JoinStr = "-"
  sb.str = "Start"
  sb.TrimBehaviour = RTrim
  sb.InjectionVariables.Add "@1", "cool"
  sb.[This is a really cool multi-line    ]
  sb.[string which can even include       ]
  sb.[symbols like " ' # ! / \ without    ]
  sb.[causing compiler errors!!           ]
  sb.[also this has @1 variable injection!]
  Test = sb.str = "Start-This is a really cool multi-line-string which can even include-symbols like "" ' # ! / \ without-causing compiler errors!!-also this has cool variable injection!"
End Function

'==============================
' PRIVATE METHODS
'==============================

Private Sub RefreshFromRaw()
  ProcessedString = StaticStringBuilder.GetString(ssb)
  
  'Replace keys
  Dim key
  If Not myInjectionVariables Is Nothing Then
    For Each key In myInjectionVariables.Keys()
      ProcessedString = VBA.Replace(ProcessedString, key, InjectionVariables(key))
    Next
  End If
  
  'No longer tainted
  Tainted = False
End Sub
Private Sub Class_Initialize()
  JoinStr = vbCrLf
End Sub

