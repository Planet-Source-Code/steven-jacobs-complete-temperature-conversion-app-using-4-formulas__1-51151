VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   5580
   ClientTop       =   1935
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   3465
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   12
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////
'Author:  Steven Jacobs
'Program Name:  Temperature Convertor"
'Date created: 01/21/2004
'All rights reserved
'////////////////////////

Option Explicit
Dim i As Integer
Dim iTmp As Integer
Dim foundField As String
Dim checkString As String
Dim currField As String
Dim whatIndex As Integer
Const alphaChars = "abcdefghijklmnopqrstuvwxyz<>,./?""'-[]{}_=+!@#$%^&*()"

Private Sub Command1_Click(Index As Integer)
If Index = 1 Then Unload Me
If Index = 2 Then refreshFields
If Index = 0 Then
If validateField Then
For i = 0 To Text1.Count - 1
If Text1(i).Text <> "" Then Call convertMe(Text1(i).Text, i): Exit For
Next i
Else
MsgBox "Input Error" & Chr(13) & Chr(13) & _
"A correct fahrenheit or celsius temperature is needed for conversion", , "Input Error"
Call resetGlobals
Exit Sub
End If
End If
End Sub

Sub resetGlobals()
foundField = ""
checkString = ""
currField = ""
End Sub

Sub convertMe(currField, whatIndex)

On Error GoTo tmpError

Dim C As Double
Dim F As Double
Dim K As Double
Dim RE As Double
Dim RA As Double
Dim tmpBase As Double

tmpBase = CDbl(currField)

'Just a double error check...not "really" needed
If currField = "" Then
MsgBox "A fahrenheit or celsius temperature is needed for conversion", , "Need a temperature"
Call resetGlobals
GoTo endMe
End If

'Error check for bounds
If (tmpBase < CDbl(-460) And whatIndex = 0) Or _
(tmpBase < CDbl(-273.15) And whatIndex = 1) Then
MsgBox "Temperature range error.  Theoretical bounds exceeded.", , "Bounds Error"
For i = 0 To Text1.Count - 1
Text1(i).Text = ""
Next i
Call resetGlobals
GoTo endMe
End If

'base calculations conversion formulas for rest of temp formulas
'///////////////////////////////////////////////////////////////
If whatIndex = 1 Then
F = ((CDbl(currField) * 9) / 5) + 32
tmpBase = tmpBase
ElseIf whatIndex = 0 Then
C = ((CDbl(currField) - 32) * 5) / 9
tmpBase = tmpBase
End If
'///////////////////////////////////////////////////////////////

'Kelvin
If whatIndex = 0 Then
K = ((tmpBase - 32) * 0.5555) + 273.15 'F
Else
K = tmpBase + 273.15 'C
End If

'Reaumur
If whatIndex = 1 Then
RE = (tmpBase * 4) / 5
ElseIf whatIndex = 0 Then
RE = ((tmpBase - 32) * 4) / 9
End If

'Rankine
If whatIndex = 0 Then
RA = tmpBase + 459.67
ElseIf whatIndex = 1 Then
RA = ((tmpBase * 1.8) + 32) + 459.67
End If

If whatIndex = 0 Then
Text1(0).Text = tmpBase
Else
Text1(0).Text = F
End If

If whatIndex = 1 Then
Text1(1).Text = tmpBase
Else
Text1(1).Text = C
End If

Text1(2).Text = RE
Text1(3).Text = RA
Text1(4).Text = K

endMe:
Exit Sub

tmpError:
MsgBox Err.Description
Call resetGlobals
GoTo endMe

End Sub

Public Function validateField() As Boolean
For i = 0 To Text1.Count - 1
If Text1(i).Text <> "" Then
iTmp = i
foundField = Text1(i).Text
Exit For
End If
Next i

'I'm aware of IsNumeric; however, doing some testing, this method proves faster.
'I did a comparison between the IsNumeric function and this method and on the
'two machines, this method provided .5 seconds faster.
For i = 1 To Len(foundField)
checkString = Mid(UCase(foundField), i, 1)
If InStr(1, UCase(alphaChars), UCase(checkString)) > 0 And _
Mid(UCase(foundField), 1, 1) <> "-" Then
Text1(iTmp).Text = ""
validateField = False
Call resetGlobals
Exit For
ElseIf foundField = "" Then
validateField = False
Call resetGlobals
Exit For
Else
validateField = True
Exit For
End If
Next i

End Function

Sub refreshFields()
For i = 0 To Text1.Count - 1
Text1(i).Text = ""
Next i
Call resetGlobals
End Sub

Private Sub Form_Load()
Call setObjectLabels
End Sub

Public Function setObjectLabels()

For i = 0 To Label1.Count - 1
Select Case CStr(i)
Case "0":
With Label1(i)
.Caption = "Temperature Conversion App"
.FontBold = True
End With
Case "1":
With Label1(i)
.Caption = "Select a Conversion Formula"
End With
Case "2":
With Label1(i)
.Caption = "Fahrenheit"
End With
 Case "3":
 With Label1(i)
.Caption = "Celsius"
End With
 Case "4":
 With Label1(i)
.Caption = "Reaumur"
End With
 Case "5":
 With Label1(i)
.Caption = "Rankine"
End With
 Case "6":
 With Label1(i)
.Caption = "Kelvin"
End With
 End Select
Next i

For i = 0 To Text1.Count - 1
Text1(i).Text = ""
Text1(i).Enabled = False
Next i

With List1
.AddItem "Fahrenheit", 0
.AddItem "Celsius", 1
End With

Command1(0).Caption = "Execute"
Command1(1).Caption = "Exit"
Command1(2).Caption = "Refresh"

Form1.Caption = "Temperature App"

End Function

Private Sub List1_Click()
For i = 0 To List1.ListCount - 1
If List1.ListIndex = i Then
Text1(i).Enabled = True
Else
Text1(i).Enabled = False
End If
Next i
End Sub
