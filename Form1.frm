VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get *.DLL Attributes"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   45
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save List"
      Height          =   420
      Left            =   45
      TabIndex        =   8
      Top             =   4680
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   690
      Left            =   45
      TabIndex        =   7
      Top             =   5130
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1217
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   540
      Width           =   6180
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Width           =   6180
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Properties"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4725
      Width           =   1860
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   6120
      TabIndex        =   3
      Top             =   405
      Width           =   6180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4995
      TabIndex        =   2
      Top             =   0
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Text            =   "C:\WINDOWS\SYSTEM32\User32.dll"
      Top             =   0
      Width           =   3885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1305
      TabIndex        =   9
      Top             =   4770
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DLL Filename:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Title

Private Sub Command1_Click()
On Error GoTo x
CommonDialog1.CancelError = True
CommonDialog1.Filter = "DLL Files | *.dll; | All Files (*.*) | *.*;"
CommonDialog1.DialogTitle = "Browse For File..."
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
x:
End Sub

Private Sub Command2_Click()
Label2.Caption = "Clearing Old Data"
Me.Refresh
If List1.ListCount > 0 Then
For i = 0 To List1.ListCount - 1
List1.RemoveItem 0
Next i
End If

Label2.Caption = "Loading DLL File"
Me.Refresh

RichTextBox1.LoadFile Text1.Text
For i = 0 To Len(Text1.Text)
If Left(Right(Text1.Text, i), 1) = "\" Then
RichTextBox1.Find Right(Text1.Text, i - 1), 0, Len(RichTextBox1.Text)
Title = Right(Text1.Text, i - 1)
GoTo p
End If
Next i
p:

Label2.Caption = "Extracting Attributes"
Me.Refresh

RichTextBox1.SelStart = RichTextBox1.SelStart + RichTextBox1.SelLength
RichTextBox1.SelLength = 0
Counter = 0
Do While j = 0
RichTextBox1.SelStart = RichTextBox1.SelStart + 1
RichTextBox1.SelLength = 1

If Asc(RichTextBox1.SelText) < 123 And Asc(RichTextBox1.SelText) > 40 Then
    Text2.Text = Text2.Text & RichTextBox1.SelText
    Counter = 0
Else
    List1.AddItem Text2.Text
    Text2.Text = ""
    Counter = Counter + 1
    If Counter = 2 Then GoTo x
End If
Loop
x:
Label2.Caption = "Complete : " & List1.ListCount & " Result(s)"
Me.Refresh
End Sub

Private Sub Command3_Click()
On Error GoTo x
CommonDialog1.CancelError = True
CommonDialog1.Filter = "Text File (*.txt) | *.txt; | All Files (*.*) | *.*;"
CommonDialog1.DialogTitle = "Save Attribute Listing..."
CommonDialog1.ShowSave
RichTextBox1.Text = ""
List1.Visible = False
For i = 0 To List1.ListCount - 1
List1.ListIndex = i
Me.Caption = "Saving... " & Int(List1.ListIndex / List1.ListCount * 100) & "% Complete"
RichTextBox1.Text = RichTextBox1.Text & vbCrLf & List1.Text
Next i
List1.ListIndex = 0
Me.Caption = "Get *.DLL Attributes"
List1.Visible = True
RichTextBox1.SaveFile CommonDialog1.FileName
x:
End Sub

Private Sub List1_Click()
Text2.Text = "Public Declare Function " & List1.Text & " Lib " & Chr(34) & Title & Chr(34) & " Alias " & Chr(34) & List1.Text & Chr(34) & " (ByVal...) As ..."
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Me.Caption = KeyAscii
End Sub
