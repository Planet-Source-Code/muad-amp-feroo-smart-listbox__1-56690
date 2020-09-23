VERSION 5.00
Begin VB.Form SMART 
   Caption         =   "SMART LISTA"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command7 
      Caption         =   "TOP"
      Height          =   400
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton Command8 
      Caption         =   "REMOVE FILTER"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5040
      TabIndex        =   12
      Top             =   4560
      Width           =   1100
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Scroll"
      Height          =   400
      Left            =   6240
      TabIndex        =   10
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Item wihout validation"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Item wihout duplicate"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FILTER"
      Height          =   400
      Left            =   3840
      TabIndex        =   5
      Top             =   4560
      Width           =   1100
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SEARCH"
      Height          =   400
      Left            =   2640
      TabIndex        =   3
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "buttom"
      Height          =   400
      Left            =   1440
      TabIndex        =   2
      Top             =   4560
      Width           =   1100
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   2370
      Left            =   2640
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "&Auto Complete"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "smart listbox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "SMART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SMART LISTBOX BY MUAD_DR


Option Compare Text
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim ADDITEM As Boolean
Dim BETA() As String
Dim SS As Integer


Private Sub Check1_Click()
Text1.SetFocus
End Sub

Private Sub Command1_Click() ' FILTER

If Text1.Text = "" Then
MsgBox StrConv("PLEASE TYPE YOUR ITEM IN TEXTBOX", vbProperCase)
Text1.SetFocus
Exit Sub
End If
Command8.Enabled = True




For I = 0 To List1.ListCount - 1
If List1.List(I) = Text1.Text Then
List2.ADDITEM List1.List(I)
End If
Next

SS = List1.ListCount
ReDim BETA(List1.ListCount)
For I = 0 To List1.ListCount - 1
BETA(I) = List1.List(I)
Next
List1.Clear

For I = 0 To List2.ListCount - 1
List1.ADDITEM List2.List(I)
Next

Me.Command1.Enabled = False
End Sub

Private Sub Command2_Click() ' go buttom

List1.ListIndex = List1.ListCount - 1

End Sub


Private Sub Command3_Click() ' find
Me.Command8.Enabled = False
Dim C As Byte
If Len(Trim(Text1.Text)) = 0 Then
MsgBox StrConv("PLEASE TYPE YOUR ITEM IN TEXTBOX", vbProperCase)
Text1.SetFocus
Exit Sub
End If
For I = 0 To List1.ListCount - 1
DoEvents
If List1.List(I) = Text1.Text Then
Sleep 500
C = C + 1
List1.ListIndex = I
End If
Next
If C > 1 Then MsgBox "DUPLICATE VALUES IN LISTA"
End Sub

Private Sub Command4_Click() 'add item to list
' WHEN ADD UNIQUE ITEM TO LISTBOX YOU MUST PREVENT CHANGE AND CLICK EVENTS
' I RECOMENDED TO REMOVE ADDITEM=TRUE
' THEN PUT DEBUG MARK AND SHOW WHAT THE CONFILICT BETWEEN ADDITEM AND CLICK AND CHANGE EVENTS

ADDITEM = True ' IMPORTANT
If Len(Trim(Text1.Text)) = 0 Then Text1.SetFocus: Exit Sub

List1.Text = Text1.Text

If List1.ListIndex > -1 Then
MsgBox vbTab & StrConv("duplicate item", vbProperCase) & vbTab
Exit Sub
End If
List1.ADDITEM Text1.Text
ADDITEM = False
End Sub

Private Sub Command5_Click() ' add duplicate item
If Len(Trim(Text1.Text)) = 0 Then Text1.SetFocus: Exit Sub



List1.ADDITEM Text1.Text
End Sub

Private Sub Command6_Click()
For Each Control In Me
If TypeOf Control Is CommandButton Then
Control.Enabled = False
End If
Next
scroll
For Each Control In Me
If TypeOf Control Is CommandButton Then
Control.Enabled = True
End If
Next

End Sub


Private Sub Command7_Click()
List1.ListIndex = 0

End Sub

Private Sub Command8_Click()

FILL_LISTA
Me.Command8.Enabled = False
Me.Command1.Enabled = True
End Sub

Private Sub Form_Load()
VB.Clipboard.Clear
Show
DoEvents

For I = 1 To 20
DoEvents
List1.ADDITEM "CAT" & I
Next
FIX_CMD
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set SMART = Nothing
End
End Sub

Private Sub List1_Click()
If ADDITEM Then Exit Sub
Text1.Text = List1.Text
End Sub

Private Sub Text1_Change()
If ADDITEM Then Exit Sub
If Check1.Value = 1 And Len(Text1.Text) >= 4 Then List1.Text = Text1.Text
End Sub



Sub scroll()
List1.ListIndex = -1
For I = 0 To List1.ListCount - 1
DoEvents
Sleep 100
List1.Text = List1.List(I)
Next
For I = List1.ListCount - 1 To 0 Step -1
DoEvents
Sleep 100
List1.Text = List1.List(I)
Next

End Sub


Friend Sub FILL_LISTA()
List2.Clear
List1.Clear

List1.Visible = False
For I = 0 To SS - 1
List1.ADDITEM BETA(I)
Next

List1.Visible = True
Erase BETA
End Sub


Friend Sub FIX_CMD()
For Each Control In Me
If TypeOf Control Is CommandButton Then
Control.Caption = "&" & StrConv(Control.Caption, vbProperCase)
End If
Next
Me.Caption = StrConv(Me.Caption, vbProperCase)
Label1.Caption = StrConv(Label1.Caption, vbProperCase)

End Sub
