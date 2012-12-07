Attribute VB_Name = "basFunction"
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : basfunction
'   �� ��  �� �� : �����Լ�
'
'   ��   ��   �� : 2007/08/20
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit


'''''gsubSS_DelRow(dsds,66)
'Public Sub gsubSS_DelRow(ss As Control, Optional R1 As Long)
'    Dim row_id As Long
'
'    If (ss.MaxRows = 0) Then
'        Call gsubSS_Clear(ss, " ", 1, 0, 1, ss.MaxCols)
'    Else
'        If (IsMissing(R1)) Then
'            row_id = ss.ActiveRow
'        Else
'            row_id = R1
'        End If
'        ss.Row = row_id
'        ss.Action = 5       'SS_ACTION_DELETE_ROW
'        If (row_id = ss.MaxRows) Then Call gsubSS_CellMove(ss, ss.MaxRows - 1, ss.Col, True)
'        ss.MaxRows = ss.MaxRows - 1
'    End If
'End Sub



'## ���߽ÿ� �������� MsgBox�� ǥ��
Public Sub DMsgBox(ByVal msg As String, ByVal title As String)

    Select Case UCase(Trim(basModule.connDB))
        Case "MIMAC"                                                '<< ��������
            MsgBox msg, vbCritical + vbOKOnly, title
        Case Else                                                   '<< ���߿�
            MsgBox msg & vbCrLf & _
                Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, title
    End Select

End Sub

'## �ѱ� MID
Public Function MidKor(ByVal vStr As String, ByVal vStart As Integer, ByVal vSize As Integer) As String
    MidKor = StrConv(MidB(StrConv(vStr, vbFromUnicode), vStart, vSize), vbUnicode)
End Function

'## �ѱ� ����
Public Function LenKor(ByVal vStr As String) As Long
    LenKor = LenB(StrConv(vStr, vbFromUnicode))
End Function

'## �������� locking
Public Sub Lock_Spread(ByRef aSprName As Object, _
                    ByVal aRowStart As Long, ByVal aRowEnd As Long, _
                    ByVal aColStart As Long, ByVal aColEnd As Long)
    
    With aSprName
        .Row = aRowStart:       .Row2 = aRowEnd
        .Col = aColStart:       .Col2 = aColEnd
        
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
End Sub

'## �������� �� : ��������
Public Sub Set_SprType_Numeric(ByRef aSprName As Object, _
                      ByVal aDecplace As Long, _
                      ByVal aMinValue As Double, ByVal aMaxValue As Double, _
                      ByVal aSepGbn As String, ByVal aValue As Double)
    
    Dim ni      As Integer
    Dim sDec    As String
    
    With aSprName
        .CellType = CellTypeNumber
        .TypeVAlign = TypeVAlignCenter
        .TypeNumberDecPlaces = aDecplace
        .TypeNumberMin = aMinValue
        .TypeNumberMax = aMaxValue
        
        If aSepGbn <> "" Then
            .TypeNumberSeparator = aSepGbn
            .TypeNumberShowSep = True
        Else
            .TypeNumberShowSep = False
        End If
        
        If aDecplace = 0 Then
            .value = Format(aValue, "#########0")
        Else
            sDec = "#########0."                            ' �Ҽ��� ǥ��
            For ni = 1 To aDecplace - 1 Step 1
                sDec = sDec & "#"
            Next ni
            sDec = sDec + "#"
            .value = Format(aValue, sDec)
        End If
    End With
End Sub

'## �������� �� : üũ�ڽ� ����
Public Sub Set_SprType_ChkBox(ByRef aSprName As Object)
    With aSprName
        .CellType = CellTypeCheckBox
        .TypeCheckCenter = True
    End With
End Sub

'## �������� �� : �ؽ�Ʈ ����
Public Sub Set_SprType_Text(ByRef aSprName As Object, _
                    ByVal aVerAlign As String, ByVal aHorAlign As String, _
                    ByVal aLength As Long, ByVal aValue As String)
    
    With aSprName
        .CellType = CellTypeEdit
        
        Select Case UCase(aVerAlign)
            Case "TOP"
                .TypeVAlign = TypeVAlignTop
            Case "BOTTOM"
                .TypeVAlign = TypeVAlignBottom
            Case "CENTER"
                .TypeVAlign = TypeVAlignCenter
        End Select
        
        Select Case UCase(aHorAlign)
            Case "LEFT"
                .TypeHAlign = TypeHAlignLeft
            Case "RIGHT"
                .TypeHAlign = TypeHAlignRight
            Case "CENTER"
                .TypeHAlign = TypeHAlignCenter
        End Select
        
        .TypeMaxEditLen = aLength
        
        .Text = aValue
    End With
End Sub

'## �������� �� : row���� n���� �׷��ؼ� ó��
Public Sub Set_SprRowBackColor_By_NRow(ByRef aSprName As Object, ByVal aStepNum As Integer, ByVal aColorColumn As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nChk        As Integer
    
    With aSprName
        If .MaxRows = 0 Then Exit Sub
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow:            .Col = 1
            .Row2 = nRow:           .Col2 = .MaxCols
            
            .BlockMode = True
                If (nRow - 1) Mod (aStepNum * 2) < aStepNum Then
                    .BackColor = basModule.GroupColor1
                    '.SelBackColor = basModule.GroupColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                    nChk = 1
                Else
                    .BackColor = basModule.BackColor2
                    '.SelBackColor = basModule.gBackColorWhite
                    .BackColorStyle = BackColorStyleUnderGrid
                    nChk = 0
                End If
            .BlockMode = False
            
            If aColorColumn <> 0 Then
                .Row = nRow
                .Col = aColorColumn
                    .CellType = CellTypeCheckBox
                    .TypeCheckCenter = True
                    .value = nChk
            End If
        Next nRow
    End With
End Sub


'## �������� �� : ���� COLUMN �׸񳢸� �׷��ؼ� ��ó��
Public Sub Set_SprRowBackColor_By_SameColor(ByRef aSprName As Object, ByVal aSprChkColumn As Long, ByVal aColorColumn As Long)
    Dim nRow        As Long
    
    Dim sTmp        As String
    Dim sComp       As String
    Dim nChkColor   As String
    Dim nColor      As Long
    
    sComp = ""
    sTmp = ""
    nChkColor = 1
    
    With aSprName
        If .MaxRows = 0 Then Exit Sub
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow:            .Col = aSprChkColumn:       sTmp = Trim(.Text)
            
            If StrComp(sComp, sTmp, vbTextCompare) <> 0 Then
                Select Case nChkColor
                    Case 0
                        nColor = basModule.BackColor2
                        nChkColor = 1
                    Case "1"
                        nColor = basModule.GroupColor1
                        nChkColor = 0
                End Select
                
                sComp = sTmp
            End If
            
            .Row = nRow:            .Col = 1
            .Row2 = nRow:           .Col2 = .MaxCols
            
            .BlockMode = True
                .BackColor = nColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            If aColorColumn <> 0 Then
                .Row = nRow
                .Col = aColorColumn
                    .CellType = CellTypeCheckBox
                    .TypeCheckCenter = True
                    Select Case nChkColor
                        Case 0
                            .value = 1
                        Case 1
                            .value = 0
                    End Select
            End If
        Next nRow
    End With
End Sub


'## ���ڿ� �߿� Ư�����ڸ� �׷��� ���ڷ� ����
Public Function Change_EspChr_To_GraphicChr(ByVal aSpcStr As String) As String
    Dim sStr        As String
    Dim sConv       As String
    
    sConv = aSpcStr
    sStr = Replace(sConv, Chr(96), Chr(-23584), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(94), Chr(-23586), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(93), Chr(-23587), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(92), Chr(-23588), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(91), Chr(-23589), 1, -1, vbTextCompare):      sConv = sStr
    
    
    sStr = Replace(sConv, Chr(64), Chr(-23616), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(63), Chr(-23617), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(59), Chr(-23621), 1, -1, vbTextCompare):      sConv = sStr
    'sStr = Replace(sConv, Chr(58), Chr(-23622), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(47), Chr(-23633), 1, -1, vbTextCompare):      sConv = sStr
    
    sStr = Replace(sConv, Chr(39), Chr(-23641), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(38), Chr(-23642), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(37), Chr(-23643), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(36), Chr(-23644), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(35), Chr(-23645), 1, -1, vbTextCompare):      sConv = sStr
    sStr = Replace(sConv, Chr(34), Chr(-23646), 1, -1, vbTextCompare):      sConv = sStr
    
    sStr = Replace(sConv, Chr(33), Chr(-23647), 1, -1, vbTextCompare):      sConv = sStr
    
    Change_EspChr_To_GraphicChr = sStr
    
End Function

'## �ؽ�Ʈ �ڽ� ������ ���콺 �˾����� �ٲ�
Public Function NoContextMenuWindowProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_CONTEXTMENU = &H7B
    
    If msg <> WM_CONTEXTMENU Then _
        NoContextMenuWindowProc = CallWindowProc(ContextMenuWindowProc, hWnd, msg, wParam, lParam)
        
End Function

'## �ؽ�Ʈ �ڽ� ������ ���콺 �˾�����
Public Sub RemoveContextMenu(ByVal text_box As TextBox)
    Const GWL_WNDPROC = (-4)
    
    ContextMenuWindowProc = SetWindowLong(text_box.hWnd, GWL_WNDPROC, AddressOf NoContextMenuWindowProc)
    
End Sub

''## Http submit ó��
'Public Function HttpRequest(ByRef aSocket As DSSocket.clsSocket, _
'                            ByVal aJobName As String, _
'                            ByRef aConditions() As String, _
'                            ByVal aPHPName As String, _
'                            ByVal aErrorDescription As String, _
'                            ByVal aMsgBoxCaption As String) As String
'
'    Dim sConditionParameter     As String
'    Dim sReceived               As String
'
'    sReceived = ""
'
'    On Error GoTo ErrorHandler
'
'    sConditionParameter = basDataTrans.Make_SendFormat(UCase(aJobName), aConditions)
'
'    If basModule.gRegID = "" Then basModule.gRegID = "0000000000000"
'    sReceived = aSocket.Submit(App.Path, basModule.LoginHost, basModule.PORT, basModule.Login_Path & aPHPName, basModule.gRegID, sConditionParameter, 0, 0, True)
'
'    If sReceived = "00000" Then   '�������
'        If basDataTrans.Format_ReceiveData(aSocket.GetBody) = False Then
'            MsgBox aErrorDescription, vbExclamation, aMsgBoxCaption
'            HttpRequest = ""
'        Else
'            HttpRequest = basDataTrans.gsRecvData
'        End If
'    Else                    '���������
'        MsgBox aSocket.GetError(sReceived), vbExclamation, aMsgBoxCaption
'        HttpRequest = ""
'    End If
'
'    Exit Function
'ErrorHandler:
'    MsgBox aErrorDescription, vbExclamation, aMsgBoxCaption
'End Function

'## ^T,^N�� �߶� aRs() 2���� �迭�� ����
Public Sub MDO(ByRef aRs() As String, ByVal aHttpResult As String, ByVal aMsgBoxCaption As String)
    Dim sRows()         As String
    Dim sCols()         As String
    Dim ni              As Long
    Dim nk              As Long

    On Error GoTo ErrorHandler
    
    sRows = Split(aHttpResult, "^N")
    sCols = Split(sRows(0), "^T")
    
    ReDim aRs(UBound(sRows), UBound(sCols))
    
    For ni = 0 To UBound(sRows) - 1
        
        sCols = Split(sRows(ni), "^T")
        For nk = 0 To UBound(sCols)
            aRs(ni, nk) = sCols(nk)
        Next
    Next
    
    Exit Sub
ErrorHandler:
    MsgBox "Error cutting function.", vbExclamation, aMsgBoxCaption
End Sub

'## Ư�� ���������� Ŭ���� �θ� �� �÷���Ŵ.
Public Sub SetSprColor1Row(ByRef aSpread As Object, ByVal ColorGbn As Integer, ByVal aRow As Long)
    Dim nRow        As Long
    
    With aSpread
        'If .Tag = aRow Then Exit Sub                '������ row�� ���� ��ĥ�� �ο� ���ٸ� �ƿ�!
        If .Tag = "" Then .Tag = "1"
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 1
            
            If ColorGbn = 1 Then
                If .BackColor <> basModule.BackColor1 Then
                    .Tag = CStr(.Row)
                    Exit For
                End If
            ElseIf ColorGbn = 2 Then
                If .BackColor <> basModule.BackColor2 Then
                    .Tag = CStr(.Row)
                    Exit For
                End If
            End If
        Next nRow
        
        If .Tag <> "" Then  '�̹� ��ĥ�� row�� �ִٸ� ������ ������� �ٲ۴�.
            .Row = CLng(.Tag):  .Row2 = CLng(.Tag)
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
            
                If ColorGbn = 1 Then
                    .BackColor = basModule.BackColor1
                ElseIf ColorGbn = 2 Then
                    .BackColor = basModule.BackColor2
                End If
                .BackColorStyle = BackColorStyleUnderGrid
                
            .BlockMode = False
        End If
        
        ' ������ row�� ������ ĥ�Ѵ�.
        .Row = aRow:        .Row2 = aRow
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            
            If ColorGbn = 1 Then
                .BackColor = basModule.SelectColor1
            ElseIf ColorGbn = 2 Then
                .BackColor = basModule.SelectColor2
            End If
            
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = aRow
    End With
End Sub

'## �������� �����ִ� �ο� Ư�� �÷����� ��ȯ
Public Function GetSprColorColValue(ByRef aSpread As Object, ByVal aCol As Long) As String
    With aSpread
        If .Tag = "" Then Exit Function
        
        .Row = .Tag
        .Col = aCol
        
        GetSprColorColValue = .Text
    End With
End Function

'## ����ȭ�� status bar text ������.
Public Sub StatusBar(ByVal aMsg As String)
    'Call DS_CLASS_MDIMAIN.txt_StatusBar(aMsg)
    
End Sub

'## ���õ� �ð�����
Public Function gGetLocalTime() As String
    Dim kLocTime    As SYSTEMTIME
    Dim sTmp        As String
    
    '�ð��� ����.
    GetLocalTime kLocTime
    
    sTmp = ""
    With kLocTime
        sTmp = sTmp & .wYear & "-"              ' �⵵
        sTmp = sTmp & .wMonth & "-"             ' ��
        sTmp = sTmp & .wDayOfWeek & "-"         ' ����(0 - 6 : �Ͽ��� 0)
        sTmp = sTmp & .wDay & "-"               ' ��
        sTmp = sTmp & .wHour & "-"              ' �ð�
        sTmp = sTmp & .wMinite & "-"            ' ��
        sTmp = sTmp & .wSecond & "-"            ' ��
        sTmp = sTmp & .wMilliseconds            ' �и���
    End With
    
    gGetLocalTime = sTmp
    
End Function




'-------------------------------------------------------------------------------------------------------------------------------
' ��´��
'-------------------------------------------------------------------------------------------------------------------------------
Sub PrintStartDoc(PaperWidth, PaperHeight, PaperSize, Orientation, TMargin, LMargin, Optional CenterOpt As Integer = 1)
    Dim psm
    Dim fsm

    On Error Resume Next
    
 ' Set the physical page size:
    PgWidth = PaperWidth                                          ' ��¼���(PageWidth)
    PgHeight = PaperHeight                                        ' ��¼���(PageHeight)
   
    Printer.Orientation = Orientation                             ' ����/�������
    Printer.PaperSize = PaperSize                                 ' ��������(A4,B4....)
    Printer.ScaleMode = vbTwips                                   ' ��������(twip: 567 = 1cm)
    
    If (CenterOpt) Then
        TBGap = (PgHeight - Printer.ScaleHeight) / 2 - TMargin '* 567  ' TOP   1cm(567 twip)���� ����
        LRGap = (PgWidth - Printer.ScaleWidth) / 2 - LMargin '* 567    ' LEFT  1cm(567 twip)���� ����
    Else
        TBGap = (PgHeight - Printer.ScaleHeight) - TMargin  '* 567  ' TOP   1cm(567 twip)���� ����
        LRGap = (PgWidth - Printer.ScaleWidth) - LMargin  '* 567    ' LEFT  1cm(567 twip)���� ����
    End If
    
    Printer.ScaleMode = psm
    sm = Printer.ScaleMode
    
    'On Error GoTo 0

End Sub

Sub PrintCurrentX(XVal)
    Printer.CurrentX = XVal - LRGap
End Sub

Sub PrintCurrentY(YVal)
    Printer.CurrentY = YVal - TBGap
End Sub

Sub PrintFontName(pFontName)
    Printer.FontName = pFontName
End Sub

Sub PrintFontSize(pSize)
    Printer.FontSize = pSize
End Sub

Sub PrinterPrint(PrintVar)
    Printer.Print PrintVar
End Sub

Sub PrintLine(bLeft0, bTop0, bLeft1, bTop1)
    Printer.Line (bLeft0 - LRGap, bTop0 - TBGap)-(bLeft1 - LRGap, bTop1 - TBGap)
End Sub

Sub PrintBox(bLeft, bTop, bWidth, bHeight)
    Printer.Line (bLeft - LRGap, bTop - TBGap)-(bLeft + bWidth - LRGap, bTop + bHeight - TBGap), , B
End Sub

Sub PrintFilledBox(bLeft, bTop, bWidth, bHeight, color)
    Printer.Line (bLeft - LRGap, bTop - TBGap)-(bLeft + bWidth - LRGap, bTop + bHeight - TBGap), color, BF
End Sub

Sub PrintCircle(bLeft, bTop, bRadius)
    Printer.Circle (bLeft - LRGap, bTop - TBGap), bRadius
End Sub
 
Sub PrintPicture(bPicture, bLeft, bTop, bWidth, bHeight)
    Printer.PaintPicture bPicture, bLeft - LRGap, bTop - TBGap, bWidth, bHeight
End Sub

Sub PrintNewPage()
    Printer.NewPage
End Sub

Sub PrintEndDoc()
    Printer.EndDoc
    Printer.ScaleMode = sm
End Sub









Public Sub GetZoom(zoomlabel As Integer)
'    'Set up the print previews zoom
'
'    Select Case zoomlabel
'        Case 0
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 200
'
'        Case 1
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 150
'
'        Case 2
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 100
'
'        Case 3
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 75
'
'        Case 4
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 50
'
'        Case 5
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 25
'
'        Case 6
'            PRT900.vaSpreadPreview1.PageViewType = 2
'            PRT900.vaSpreadPreview1.PageViewPercentage = 10
'
'        Case 7
'            PRT900.vaSpreadPreview1.PageViewType = 3
'
'        Case 8
'            PRT900.vaSpreadPreview1.PageViewType = 4
'
'        Case 9
'            PRT900.vaSpreadPreview1.PageViewType = 0
'
'        Case 10
'            PRT900.vaSpreadPreview1.PageViewType = 5
'            PRT900.vaSpreadPreview1.PageMultiCntH = 2
'            PRT900.vaSpreadPreview1.PageMultiCntV = 1
'
'        Case 11
'            PRT900.vaSpreadPreview1.PageViewType = 5
'            PRT900.vaSpreadPreview1.PageMultiCntH = 3
'            PRT900.vaSpreadPreview1.PageMultiCntV = 1
'
'        Case 12
'            PRT900.vaSpreadPreview1.PageViewType = 5
'            PRT900.vaSpreadPreview1.PageMultiCntH = 2
'            PRT900.vaSpreadPreview1.PageMultiCntV = 2
'
'        Case 13
'            PRT900.vaSpreadPreview1.PageViewType = 5
'            PRT900.vaSpreadPreview1.PageMultiCntH = 3
'            PRT900.vaSpreadPreview1.PageMultiCntV = 2
'
'    End Select
End Sub
