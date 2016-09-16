Attribute VB_Name = "mdlModule"
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type FCS_Node
    Father As Long
    Name As String
    LinkTo As Long
    NextN As Long
End Type

Public Type HL_Info
    KeyWord As String
    R As Byte
    G As Byte
    B As Byte
End Type

Public FCS_Tree() As FCS_Node
Public HL_List() As HL_Info
Public PoundC As Long, SlashC As Long

Public Function InitializeTree(FilePath As String) As Long
ReDim FCS_Tree(0) As FCS_Node
Open FilePath For Input As #1
Dim TempStr As String, FatherNode As Long, LstNode As Long

Do Until EOF(1)
    Line Input #1, TempStr
    Select Case TempStr
        Case "->"
        FatherNode = UBound(FCS_Tree)
        Case "<-"
        FatherNode = FCS_Tree(FatherNode).Father
        Case Else
        ReDim Preserve FCS_Tree(UBound(FCS_Tree) + 1) As FCS_Node
        FCS_Tree(UBound(FCS_Tree)).Name = TempStr
        FCS_Tree(UBound(FCS_Tree)).NextN = FCS_Tree(FatherNode).LinkTo
        FCS_Tree(UBound(FCS_Tree)).Father = FatherNode
        FCS_Tree(FatherNode).LinkTo = UBound(FCS_Tree)
    End Select
Loop

Close #1
End Function
Public Function ReturnID(Expression As String, FatherNode As Long) As Long
Dim NodeNow As Long
NodeNow = FCS_Tree(FatherNode).LinkTo
Do Until NodeNow = 0
    If FCS_Tree(NodeNow).Name = Expression Then
        ReturnID = NodeNow
        Exit Function
    End If
    NodeNow = FCS_Tree(NodeNow).NextN
Loop
ReturnID = -1
End Function
Public Function FindHeadOfName(Expression As String, FatherNode As Long, CharFJoin As String) As String
Dim NodeNow As Long
NodeNow = FCS_Tree(FatherNode).LinkTo
Do Until NodeNow = 0
    If InStr(1, FCS_Tree(NodeNow).Name, Expression) = 1 Then FindHeadOfName = FindHeadOfName & NodeNow & CharFJoin
    NodeNow = FCS_Tree(NodeNow).NextN
Loop
If FindHeadOfName = "" Then FindHeadOfName = -1 'return for failure
End Function

'Public Function DebugOptDict() 'Run a scan for forward chain star
'For i = 0 To UBound(FCS_Tree)
'    Debug.Print "name: " & FCS_Tree(i).Name & " link to: " & FCS_Tree(i).LinkTo & " next node: " & FCS_Tree(i).NextN
'Next
'End Function

'Public Function DebugOptHL() 'Run a scan for list
'For i = 0 To UBound(HL_List)
'    With HL_List(i)
'    Debug.Print "name: " & .KeyWord & " r: " & .R & " g: " & .G & " B: " & .B
'    End With
'Next
'End Function

Public Function LoadHL(FilePath As String) As Long
Open FilePath For Input As #1
Dim TempVar As Variant, DummyInfo As HL_Info
ReDim HL_List(0) As HL_Info
Do Until EOF(1)
    Line Input #1, TempVar
    Select Case TempVar
        Case "->"
        Line Input #1, TempVar
        TempVar = Split(TempVar, " ")
        DummyInfo.R = TempVar(0)
        DummyInfo.G = TempVar(1)
        DummyInfo.B = TempVar(2)
        Case Else
        Select Case Left(TempVar, 2)
            Case "#!"
            TempVar = Split(TempVar, " ")
            PoundC = RGB(TempVar(1), TempVar(2), TempVar(3))
            Case "//"
            TempVar = Split(TempVar, " ")
            SlashC = RGB(TempVar(1), TempVar(2), TempVar(3))
            Case Else
            ReDim Preserve HL_List(UBound(HL_List) + 1) As HL_Info
            HL_List(UBound(HL_List)) = DummyInfo
            HL_List(UBound(HL_List)).KeyWord = TempVar
        End Select
    End Select
Loop
Close #1
End Function

Public Function GetFile() As String
Dim i As Integer
Dim kuang As OPENFILENAME
Dim filename As String
kuang.lStructSize = Len(kuang)
kuang.hwndOwner = frmMain.hWnd
kuang.hInstance = App.hInstance
kuang.lpstrFile = Space(254)
kuang.nMaxFile = 255
kuang.lpstrFileTitle = Space(254)
kuang.nMaxFileTitle = 255
kuang.lpstrInitialDir = App.Path
kuang.flags = 6148
'过虑对话框文件类型
kuang.lpstrFilter = "所有文件 (*.*)" + Chr$(0) + "*.*" + Chr$(0) '"文本文件 (*.TXT)" + Chr$(0) + "*.TXT" + Chr$(0) + "所有文件 (*.*)" + Chr$(0) + "*.*" + Chr$(0)
'对话框标题栏文字
kuang.lpstrTitle = "保存文件的路径及文件名..."
i = GetSaveFileName(kuang) '显示保存文件对话框
If i >= 1 Then '取得对话中用户选择输入的文件名及路径
    filename = kuang.lpstrFile
    filename = Left(filename, InStr(filename, Chr(0)) - 1)
End If
GetFile = filename
End Function
