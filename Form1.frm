VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Undo'n'Redo By Joe"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbmain 
      Height          =   2550
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4498
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redo"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Undo"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "If you like this please vote"
      Height          =   270
      Left            =   2640
      TabIndex        =   3
      Top             =   180
      Width           =   1965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is created By Joe Pritoulski
'This code is a complete undo-redo routine
'with no bugs so far...
'If you find any tell me: ICQ:77803604 e-mail:joe_@mail333.com
'
'This code is a freewear, use it as much as you like...
'
'
'
'
'
'
' Max undo steps + 1 (note you can use from 2 to 10000 steps 
' depending on machine capabilities and textbox usage, recomended 1000 couse on my machine 
' shows "Out of memory" worning if maxundo>1000
Const maxundo = 11 
Dim uInd As Integer 'Tells how many undo or redo steps to take
Dim uSav As Boolean 'Cancels memorising string while undo or redo
Dim uStp As Integer 'Tells (just for bug fixing) when to stop undo
Dim rtbu1(maxundo) As String ' Buffer string

Private Sub Command1_Click()
If Len(rtbmain.Text) = 1 Then Command1.Enabled = False 'Check if the textbox is empty to stop undo
uSav = False 
Command2.Enabled = True
uInd = uInd - 1
If uInd < 1 Then uInd = maxundo
rtbmain.TextRTF = rtbu1(uInd) 'this line is for Richtextbox for, simple textbox replace with rtbmain.Text = rtbu1(uInd)
uStp = uStp - 1
If uStp <= 1 Then Command1.Enabled = False
End Sub

Private Sub Command2_Click()
uSav = False
Command1.Enabled = True
uInd = uInd + 1
If uInd > maxundo Then uInd = 1
rtbmain.TextRTF = rtbu1(uInd) 'this line is for Richtextbox for, simple textbox replace with rtbmain.Text = rtbu1(uInd)
uStp = uStp + 1
If uStp = maxundo Then Command2.Enabled = False
End Sub

Private Sub Form_Load()
uSav = False
uStp = maxundo
uInd = 0
cu = 0
Command1.Enabled = False
Command2.Enabled = False
End Sub

Private Sub rtbmain_Change()
If uSav = True Then
uInd = uInd + 1
If uInd = (maxundo + 1) Then uInd = 1
rtbu1(uInd) = rtbmain.TextRTF 'this line is for Richtextbox for, simple textbox replace with rtbu1(uInd) = rtbmain.Text
End If
End Sub

Private Sub rtbmain_KeyDown(KeyCode As Integer, Shift As Integer)
uSav = True
uStp = maxundo
Command1.Enabled = True
Command2.Enabled = False
End Sub
