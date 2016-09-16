VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   4710
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   5910
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrRefresh 
      Interval        =   55
      Left            =   5520
      Top             =   0
   End
   Begin VB.ListBox lstDict 
      Appearance      =   0  'Flat
      Height          =   1650
      ItemData        =   "frmMain.frx":0000
      Left            =   240
      List            =   "frmMain.frx":0002
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox txtText 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      RightMargin     =   1.22222e7
      TextRTF         =   $"frmMain.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblState 
      Caption         =   "Ready......"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   5895
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "打开"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "保存"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPart As String, WordPointer As Long, WordPointer2 As Long
Dim FP As String, firstProcess As Boolean
Private Sub Form_Load()
lblState.Caption = "Loading Statements......"
InitializeTree App.Path & "\Statements"
lblState.Caption = "Ready......"
LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
End Sub

Private Sub Form_Resize()
txtText.Width = Me.Width - 220
If Me.Height - 900 - lblState.Height > 0 Then txtText.Height = Me.Height - 900 - lblState.Height
lblState.Top = txtText.Height
lblState.Width = Screen.Width
End Sub

Private Sub lstDict_GotFocus()
txtText.SetFocus
End Sub

Private Sub mnuOpen_Click()
txtText.Text = ""
FP = GetFile()
txtText.LoadFile FP
Me.Caption = FP
End Sub

Private Sub mnuSave_Click()
If Me.Caption = "" Then FP = GetFile()
txtText.SaveFile FP, 1
End Sub

Private Sub tmrRefresh_Timer()
lstDict.Left = GetCurPos(txtText).x * 15 + 255
lstDict.Top = GetCurPos(txtText).y * 15 + 255
End Sub

'Private Sub txtText_Change()
'On Error Resume Next
'If txtText.Text = "" Then Exit Sub
'Dim x As Long, y As Long, PrevPoint As Long, varSplit As Variant
'
'x = GetCurPos2(txtText).x
'y = GetCurPos2(txtText).y
'PrevPoint = txtText.SelStart
'varSplit = Split(txtText.Text, Chr(10))
'txtText.SelStart = txtText.SelStart - x + 1
'txtText.SelLength = Len(varSplit(y - 1))
'txtText.SelColor = vbBlack
'
'Dim lngC As Long, CheckPoint As Long, StartPoint As Long
'StartPoint = txtText.SelStart
'CheckPoint = 1
'
'For lngC = 1 To UBound(HL_List)
'    Do Until InStr(CheckPoint, varSplit(y - 1), HL_List(lngC).KeyWord) = 0
'        CheckPoint = InStr(CheckPoint, varSplit(y - 1), HL_List(lngC).KeyWord)
'        txtText.SelStart = CheckPoint + StartPoint - 1
'        If txtText.SelStart <> 0 Then txtText.SelStart = txtText.SelStart - 1
'        txtText.SelLength = Len(HL_List(lngC).KeyWord) + 2
'        txtText.SelColor = RGB(HL_List(lngC).R, HL_List(lngC).G, HL_List(lngC).B)
'        CheckPoint = CheckPoint + Len(HL_List(lngC).KeyWord)
'        txtText.SelStart = CheckPoint + StartPoint
'        txtText.SelColor = vbBlack
'    Loop
'    CheckPoint = 1
'    Next
'
'If InStr(1, varSplit(y - 1), "#") <> 0 Then
'    txtText.SelStart = StartPoint - 1 + InStr(1, varSplit(y - 1), "#")
'    txtText.SelLength = Len(varSplit(y - 1)) - InStr(1, varSplit(y - 1), "#") + 1
'    txtText.SelColor = PoundC
'    txtText.SelStart = StartPoint + Len(varSplit(y - 1))
'    txtText.SelColor = vbBlack
'End If
'
'If InStr(1, varSplit(y - 1), "//") <> 0 Then
'    txtText.SelStart = StartPoint - 1 + InStr(1, varSplit(y - 1), "#")
'    txtText.SelLength = Len(varSplit(y - 1)) - InStr(1, varSplit(y - 1), "#") + 2
'    txtText.SelColor = SlashC
'    txtText.SelStart = StartPoint + Len(varSplit(y - 1))
'    txtText.SelColor = vbBlack
'End If
'
'firstProcess = False
'txtText.SelStart = PrevPoint
'txtText.SelLength = 0
'txtText.SelColor = vbBlack
'End Sub


Private Sub txtText_KeyPress(KeyAscii As Integer) 'For getting words

Select Case KeyAscii
    Case 8 'Backspace
    If strPart <> "" Then
    lblState.Caption = "Now typing: " & strPart
    Else:
        lblState.Caption = "Ready......"
        WordPointer = 0
        WordPointer2 = 0
    End If
    If strPart <> "" Then strPart = Left(strPart, Len(strPart) - 1)
    LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
    Case 32 'Space
    If strPart <> "" Then
        If FindHeadOfName(strPart, WordPointer, "*") <> "-1" Then
            txtText.SelText = Right(lstDict.List(lstDict.ListIndex), Len(lstDict.List(lstDict.ListIndex)) - Len(strPart))
            lblState.Caption = "Filled!"
            KeyAscii = 0
            WordPointer = 0
            WordPointer2 = ReturnID(lstDict.List(lstDict.ListIndex), WordPointer)
            strPart = ""
        Else
            strPart = ""
        End If
        
    Else
        
        WordPointer = 0
        WordPointer2 = 0
        LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
    End If
    Case 13 'Enter
        WordPointer = 0
        WordPointer2 = 0
        strPart = ""
        LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
        lblState.Caption = "Ready......"
    Case 46 '.
    If WordPointer2 <> 0 Then WordPointer = WordPointer2
    LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
    Case 60
    KeyAscii = 0
    txtText.SelText = "<>"
    txtText.SelText = "<>"
    txtText.SelStart = txtText.SelStart - 1
    Case 40
    KeyAscii = 0
    txtText.SelText = "()"
    txtText.SelStart = txtText.SelStart - 1
    Case Else
    strPart = strPart & Chr(KeyAscii)
    LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
    lblState.Caption = "Now typing: " & strPart
    LoadListFromString FindHeadOfName(strPart, WordPointer, "*"), "*"
    If strPart = lstDict.List(lstDict.ListIndex) Then
        WordPointer2 = ReturnID(lstDict.List(lstDict.ListIndex), WordPointer)
        strPart = ""
    End If
End Select
lstDict.Left = GetCurPos(txtText).x * 15 + 255
lstDict.Top = GetCurPos(txtText).y * 15 + 255
End Sub

Private Sub LoadListFromString(Expression As String, CharFSplit As String)
lstDict.Clear
Dim varSplit As Variant, intC As Long, MaxLen As Long
varSplit = Split(Expression, CharFSplit)
For intC = 0 To UBound(varSplit) - 1
    lstDict.AddItem FCS_Tree(varSplit(intC)).Name
Next
If Expression <> "-1" Then lstDict.ListIndex = 0
End Sub
