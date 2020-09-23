VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Undo Redo example:1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Example 2"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Redo"
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Undo"
      Height          =   255
      Left            =   6720
      TabIndex        =   1
      Top             =   4680
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8070
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
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
Form2.Show
End Sub
Private Sub Form_Load()
' set var as new class
Set undoredo = New undoredo
' set class attributes
Set undoredo.rtf1 = Me.rtf1 ' you must set this to an richtextbox control on your form
undoredo.maxundo = 30 ' defaults to 100 so you don't need to set this
' load text from file
rtf1.LoadFile "example1.rtf"
undoredo.pushstack ' call pushstack so we can undo back to original text
End Sub

Private Sub rtf1_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode
' Now all we have to do is call pushstack when we want to remember a change
' in this example the stack is only pushed when the space or return key is pressed

'NOTE: since the richtextbox is updated by undo and redo do not use the
' richtextbox.change event to call pushstack selchange is ok / see example 2

If KeyCode = 13 Or KeyCode = 32 Then
   undoredo.pushstack
End If
End Sub


