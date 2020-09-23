VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form2 
   Caption         =   "Undo Redo Example:2"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Insert some text"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Undo"
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redo"
      Height          =   255
      Left            =   7560
      TabIndex        =   0
      Top             =   4680
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8070
      _Version        =   393217
      TextRTF         =   $"Form2.frx":0000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dim var as class
Dim undoredo As undoredo

Private Sub Command1_Click()
undoredo.undo
End Sub

Private Sub Command2_Click()
undoredo.redo
End Sub

Private Sub Command3_Click()
undoredo.pushstack
rtf1.SelRTF = "some TEXT"
undoredo.pushstack
End Sub
Private Sub Form_Load()
Set undoredo = New undoredo
Set undoredo.rtf1 = Me.rtf1 ' you must set this to an richtextbox control on your form
rtf1.LoadFile "example2.rtf"
undoredo.pushstack ' call pushstack so we can undo back to original text
End Sub

Private Sub rtf1_KeyUp(KeyCode As Integer, Shift As Integer)
undoredo.pushstack
End Sub
