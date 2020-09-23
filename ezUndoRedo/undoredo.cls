VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "undoredo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarrtf1 As RichTextBox 'local copy
'local variable(s) to hold property value(s)
Private mvarstack As Collection 'local copy
'local variable(s) to hold property value(s)
Private mvarstackpos As Long 'local copy
'local variable(s) to hold property value(s)
Public maxundo As Long
'local variable(s) to hold property value(s)
Private mvarselectstart As Collection 'local copy
Public Property Set selectstart(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.selectstart = Form1
    Set mvarselectstart = vData
End Property

Public Property Get selectstart() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.selectstart
    Set selectstart = mvarselectstart
End Property

Public Property Let stackpos(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.stackpos = 5
    mvarstackpos = vData
End Property
Public Property Get stackpos() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.stackpos
    stackpos = mvarstackpos
End Property
Public Sub pushstack()
On Error Resume Next
While stack.Count > Me.maxundo
 Me.stack.Remove (1)
 Me.selectstart.Remove (1)
Wend
Me.stackpos = Me.stack.Count + 1
Me.selectstart.Add rtf1.SelStart
Me.stack.Add Me.rtf1.TextRTF
'Debug.Print Me.stack.Count
End Sub
Public Property Set stack(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.stack = Form1
    Set mvarstack = vData
End Property
Public Property Get stack() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.stack
    Set stack = mvarstack
End Property
Public Sub undo()
On Error GoTo bail
Me.rtf1.TextRTF = Me.stack.Item(Me.stackpos - 1)
Me.rtf1.SelStart = Me.selectstart.Item(Me.stackpos - 1)
Me.stackpos = Me.stackpos - 1
rtf1.SetFocus
'Debug.Print Me.stackpos
Exit Sub
bail:
End Sub
Public Sub redo()
On Error GoTo bail
Me.rtf1.TextRTF = Me.stack.Item(Me.stackpos + 1)
Me.rtf1.SelStart = Me.selectstart.Item(Me.stackpos + 1)
Me.stackpos = Me.stackpos + 1
rtf1.SetFocus
'Debug.Print Me.stackpos
Exit Sub
bail:
End Sub
Public Property Set rtf1(ByVal vData As RichTextBox)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.rtf1 = Form1
    Set mvarrtf1 = vData
    Set Me.stack = New Collection
    Set Me.selectstart = New Collection
End Property
Public Property Get rtf1() As RichTextBox
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.rtf1
    Set rtf1 = mvarrtf1
End Property

Private Sub Class_Initialize()
' set defaults
Me.maxundo = 100
End Sub
