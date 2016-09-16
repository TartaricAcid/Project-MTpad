Attribute VB_Name = "mdlCursor"
    Public Const WM_USER = &H400
    Public Const EM_EXGETSEL = WM_USER + 52
    Public Const EM_LINEFROMCHAR = &HC9
    Public Const EM_LINEINDEX = &HBB
    Public Const EM_GETSEL = &HB0
    Public Type CHARRANGE
     cpMin As Long
     cpMax As Long
    End Type
    Public Type POINTAPI
     x As Long
     y As Long
    End Type
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

    'ȡ�ù�����ڵ��к���
    Public Function GetCurPos(ByRef TextControl As Control) As POINTAPI
    GetCaretPos GetCurPos
    End Function
    
    'ȡ�ù�����ڵ��к���
    Public Function GetCurPos2(ByRef TextControl As Control) As POINTAPI
     Dim LineIndex As Long
     Dim SelRange As CHARRANGE
     Dim TempStr As String
     Dim TempArray() As Byte
     Dim CurRow As Long
     Dim CurPos As POINTAPI
     TempArray = StrConv(TextControl.Text, vbFromUnicode)
     'ȡ�õ�ǰ��ѡ���ı���λ�� ������ RichTextBox
     'TextControl �� EM_GETSEL ��Ϣ
     Call SendMessage(TextControl.hWnd, EM_EXGETSEL, 0, SelRange)
     '���ݲ���wParamָ�����ַ�λ�÷��ظ��ַ����ڵ��к�
     CurRow = SendMessage(TextControl.hWnd, EM_LINEFROMCHAR, SelRange.cpMin, 0)
     'ȡ��ָ���е�һ���ַ���λ��
     LineIndex = SendMessage(TextControl.hWnd, EM_LINEINDEX, CurRow, 0)
     If SelRange.cpMin = LineIndex Then
     GetCurPos2.x = 1
     Else
     TempStr = String(SelRange.cpMin - LineIndex, 13)
     '���Ƶ�ǰ�п�ʼ��ѡ���ı���ʼ���ı�
     CopyMemory ByVal StrPtr(TempStr), ByVal StrPtr(TempArray) + LineIndex, SelRange.cpMin - LineIndex
     TempArray = TempStr
     'ɾ�����õ���Ϣ
     ReDim Preserve TempArray(SelRange.cpMin - LineIndex - 1)
     'ת��Ϊ Unicode
     TempStr = StrConv(TempArray, vbUnicode)
     GetCurPos2.x = Len(TempStr) + 1
     End If
     GetCurPos2.y = CurRow + 1
    End Function

