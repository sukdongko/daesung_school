VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MTX010 
   Caption         =   "�ð�ǥ ����� >> ������ �ð�ǥ ���"
   ClientHeight    =   9315
   ClientLeft      =   900
   ClientTop       =   2955
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   17625
   Begin MSComctlLib.ImageList imgTrx 
      Left            =   8790
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MTX010.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MTX010.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame3"
      Height          =   735
      Left            =   30
      TabIndex        =   23
      Top             =   30
      Width           =   15405
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   675
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   15345
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3300
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   1
            Top             =   165
            Width           =   1035
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "��ü ������ ����"
            Height          =   465
            Left            =   7950
            TabIndex        =   3
            Top             =   90
            Width           =   1965
         End
         Begin VB.CommandButton cmdFindMtx 
            Caption         =   "��ȸ�ϱ�"
            Height          =   450
            Left            =   360
            TabIndex        =   0
            Top             =   90
            Width           =   1500
         End
         Begin VB.ComboBox cboLsnType 
            Height          =   300
            Left            =   5580
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   2
            Top             =   172
            Width           =   2235
         End
         Begin VB.Label Label6 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   2190
            TabIndex        =   31
            Top             =   217
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� ����"
            Height          =   210
            Left            =   4350
            TabIndex        =   25
            Top             =   217
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraTrx 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame1"
      Height          =   2025
      Left            =   3600
      TabIndex        =   17
      Top             =   9930
      Width           =   5085
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   1965
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   5025
         Begin VB.CommandButton cmdControlDeleteTrx 
            Caption         =   "�� ��"
            Height          =   400
            Left            =   3570
            TabIndex        =   16
            Top             =   1410
            Width           =   1100
         End
         Begin VB.TextBox txtControlTrxCD 
            Enabled         =   0   'False
            Height          =   375
            Left            =   3600
            TabIndex        =   12
            Text            =   "txtControlTrxCD"
            Top             =   120
            Width           =   585
         End
         Begin VB.CommandButton cmdControlUpdateTrx 
            Caption         =   "�� ��"
            Height          =   400
            Left            =   2280
            TabIndex        =   15
            Top             =   1410
            Width           =   1100
         End
         Begin VB.CommandButton cmdControlInsertTrx 
            Caption         =   "�� ��"
            Height          =   400
            Left            =   990
            TabIndex        =   14
            Top             =   1410
            Width           =   1100
         End
         Begin VB.TextBox txtControlTrxNM 
            Height          =   375
            Left            =   1020
            MaxLength       =   100
            TabIndex        =   11
            Text            =   "txtControlTrxNM"
            Top             =   120
            Width           =   2565
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "����� ���볻�븸 �����մϴ�."
            Height          =   255
            Left            =   2190
            TabIndex        =   30
            Top             =   930
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   -120
            TabIndex        =   27
            Top             =   870
            Width           =   975
         End
         Begin VB.Label lblControlTrxColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"MTX010.frx":08A4
            Height          =   795
            Left            =   1020
            TabIndex        =   13
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ð�����"
            Height          =   210
            Left            =   -90
            TabIndex        =   22
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame3"
      Height          =   8505
      Left            =   9600
      TabIndex        =   19
      Top             =   840
      Width           =   5835
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   8445
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   5775
         Begin VB.Frame Frame6 
            Height          =   30
            Left            =   270
            TabIndex        =   29
            Top             =   1620
            Width           =   5295
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   1200
            Top             =   -150
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00F7EFE7&
            BorderStyle     =   0  '����
            Height          =   1995
            Left            =   660
            TabIndex        =   26
            Top             =   1590
            Width           =   4785
            Begin VB.TextBox txtTrxCD 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1830
               TabIndex        =   8
               Text            =   "txtTrxCD"
               Top             =   1470
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtTrxNM 
               Enabled         =   0   'False
               Height          =   300
               Left            =   150
               TabIndex        =   7
               Text            =   "txtTrxNM"
               Top             =   1470
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.CommandButton cmdTrxSel 
               Caption         =   "�ð�ǥ ���"
               Height          =   1155
               Left            =   2730
               Picture         =   "MTX010.frx":08BA
               Style           =   1  '�׷���
               TabIndex        =   9
               Top             =   330
               Width           =   2000
            End
            Begin VB.Label lblTrxStatus 
               BackStyle       =   0  '����
               Caption         =   "lblTrxStatus"
               Height          =   825
               Left            =   60
               TabIndex        =   28
               Top             =   690
               Width           =   2445
            End
         End
         Begin VB.ComboBox cboTrx 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   1650
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   4
            Top             =   390
            Width           =   1875
         End
         Begin VB.CommandButton cmdControlTrx 
            Caption         =   "�ð����� ����"
            Height          =   465
            Left            =   1650
            TabIndex        =   6
            Top             =   870
            Width           =   1635
         End
         Begin VB.Label lblTrxColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"MTX010.frx":0CFC
            Height          =   795
            Left            =   3780
            TabIndex        =   5
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "������ �ð�����"
            Height          =   210
            Left            =   -210
            TabIndex        =   21
            Top             =   435
            Width           =   1635
         End
      End
   End
   Begin FPSpread.vaSpread sprTrx 
      Height          =   8505
      Left            =   60
      TabIndex        =   10
      Top             =   840
      Width           =   9435
      _Version        =   393216
      _ExtentX        =   16642
      _ExtentY        =   15002
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   10
      ScrollBars      =   0
      SpreadDesigner  =   "MTX010.frx":0D12
   End
End
Attribute VB_Name = "MTX010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : MTX010
'   �� ��  �� �� : �ð�ǥ ����� >> ������ �ð�ǥ ���
'
'   ��   ��   �� : 2007/10/29
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit


Private sKaeyol     As String


Private Sub Form_Click()
    fraTrx.Visible = False
End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0, 15700, 9980
    
    Me.Tag = "LOAD"
        With sprTrx
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With cboLsnType
            .Clear
            .AddItem "A type" & Space(30) & "A"
            .AddItem "B type" & Space(30) & "B"
            .AddItem "C type" & Space(30) & "C"
            
            .ListIndex = 0
        End With
        
        With cboLsnType
            .Clear
            .AddItem "A type" & Space(30) & "A"
            .AddItem "B type" & Space(30) & "B"
            .AddItem "C type" & Space(30) & "C"
            
            .ListIndex = 0
        End With
        
        fraTrx.ZOrder 0
        fraTrx.Visible = False
        
        Call init_Form
        
    Me.Tag = ""
    
    Call cboTrx_Click
    
End Sub

Private Sub init_Form()
    
    chkAll.Value = 0
    
    txtControlTrxNM.Text = ""
    lblTrxColor.Caption = "Ŭ�� ��" & vbCrLf & "���� ����"
    txtTrxCD.Text = ""
    txtTrxNM.Text = ""
    
    
    cmdTrxSel.Caption = "�ð�ǥ �����ϱ�"
    cmdTrxSel.Tag = "SELECT"
    Set cmdTrxSel.Picture = imgTrx.ListImages.Item(2).Picture
    
    lblTrxStatus.Caption = "�ð�ǥ �����ϱ�" & vbCrLf & "�����Դϴ�." ' & vbCrLf & cmdTrxSel.Tag
    lblTrxStatus.ForeColor = basModule.SectionColor2
    lblTrxStatus.FontBold = True
    lblTrxStatus.Font.Size = 12
    
    sKaeyol = ""
    
End Sub

Private Sub cmdTrxSel_Click()
    Select Case cmdTrxSel.Tag
        Case "SELECT"
                        
            cmdTrxSel.Caption = "���ýð�ǥ �����ϱ�"
            cmdTrxSel.Tag = "DELETE"
            Set cmdTrxSel.Picture = imgTrx.ListImages.Item(1).Picture
            
            lblTrxStatus.Caption = "���� �ð�ǥ" & vbCrLf & "�����ϱ� �����Դϴ�." ' & vbCrLf & cmdTrxSel.Tag
            lblTrxStatus.ForeColor = basModule.SectionColor1
            lblTrxStatus.FontBold = False
            lblTrxStatus.FontItalic = True
            lblTrxStatus.Font.Size = 12
            
        Case "DELETE"
            
            cmdTrxSel.Caption = "�ð�ǥ �����ϱ�"
            cmdTrxSel.Tag = "SELECT"
            Set cmdTrxSel.Picture = imgTrx.ListImages.Item(2).Picture
            
            lblTrxStatus.Caption = "�ð�ǥ �����ϱ�" & vbCrLf & "�����Դϴ�." ' & vbCrLf & cmdTrxSel.Tag
            lblTrxStatus.ForeColor = basModule.SectionColor2
            lblTrxStatus.FontBold = True
            lblTrxStatus.FontItalic = False
            lblTrxStatus.Font.Size = 12
            
    End Select
End Sub


'###############################################################################################################################
'###############################################################################################################################





Private Sub cmdFindMtx_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim sTrxCD      As String
    Dim sTrx        As String
    Dim nLesson     As Integer
    Dim nWeeks      As Integer
    Dim nColor      As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    On Error GoTo ErrStmt
    
    sKaeyol = Trim(Right(cboKaeyol.Text, 30))       '< 2007.12.18 : �迭�߰�
    
    '<< �ʱ�ȭ
    With sprTrx
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow
        
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With
    
    
    sStr = ""
    sStr = sStr & "  SELECT A.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, A.TRX_CL "
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   Where A.ACID   = B.ACID"
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                                '< 2007.12.18 : �迭�߰�
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & sKaeyol & "'"                       '< 2007.12.18 : �迭�߰�
    If chkAll.Value = 0 Then
        sStr = sStr & " AND A.TRXCD  LIKE '" & Trim(Right(cboLsnType.Text, 30)) & "%'"
    End If
    sStr = sStr & "  UNION ALL"
    sStr = sStr & "  SELECT A.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, A.TRX_CL"
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   Where A.ACID   = B.ACID"
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                                '< 2007.12.18 : �迭�߰�
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & sKaeyol & "'"                       '< 2007.12.18 : �迭�߰�
    sStr = sStr & "     AND A.TRXCD LIKE 'PB%'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' LSNTYPE
'        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                nColor = 0
                sTrx = ""
'                    If IsNull(.Fields("TRXCD")) = False Then
'                         sTrx = Trim(.Fields("TRXCD"))
'                         sTrxCD = sTrx
'                    End If
                                    
'                                If IsNull(.Fields("TRXNM")) = False Then sTrx = Trim(.Fields("TRXNM")) & Space(30) & sTrx
                                If IsNull(.Fields("TRXNM")) = False Then sTrx = Trim(.Fields("TRXNM"))
                                
                nLesson = 0:    If IsNull(.Fields("LESSON")) = False Then nLesson = CInt(.Fields("LESSON"))
                nWeeks = 0:     If IsNull(.Fields("WEEKS")) = False Then nWeeks = CInt(.Fields("WEEKS"))
                nColor = 0:     If IsNull(.Fields("TRX_CL")) = False Then nColor = CLng(.Fields("TRX_CL"))
                
                Select Case nWeeks
                    Case 2
                        sprTrx.Col = 1
                    Case 3
                        sprTrx.Col = 2
                    Case 4
                        sprTrx.Col = 3
                    Case 5
                        sprTrx.Col = 4
                    Case 6
                        sprTrx.Col = 5
                    Case 7
                        sprTrx.Col = 6
                    Case 1
                        sprTrx.Col = 7
                End Select
                sprTrx.Row = nLesson
                sTmp = sprTrx.Text
                    If InStr(1, sTmp, sTrx, vbTextCompare) = 0 Then
                        If basFunction.LenKor(sTmp) > 0 Then
                            sTrx = sTmp & vbCrLf & sTrx
                        End If
                        Call basFunction.Set_SprType_Text(sprTrx, "TOP", "LEFT", basFunction.LenKor(sTrx), Trim(sTrx))
                        sprTrx.TypeEditMultiLine = True
                    End If
                
                sprTrx.Row2 = sprTrx.Row
                sprTrx.Col2 = sprTrx.Col
                sprTrx.BlockMode = True
                    sprTrx.BackColor = nColor
                    sprTrx.BackColorStyle = BackColorStyleUnderGrid
                sprTrx.BlockMode = False
                
                
                .MoveNext
            Next nRec
        End If
    End With

    MsgBox "������ �ð�ǥ ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"
End Sub






















'###############################################################################################################################
'###############################################################################################################################









'>> ������ �ð����� ����
Private Sub cmdControlDeleteTrx_Click()
    Dim DBCmd       As ADODB.Command            '<< �л� �� ���� ����ϱ�
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nEXE        As Long
    Dim ni          As Long
    Dim nRec        As Long
    
    Dim sTrxCD      As String
    Dim nindex      As Integer
    Dim sGbnTrx     As String
    
    Dim sDiv()      As String
    
    On Error GoTo ErrStmt
    
    Select Case Left(Trim(txtControlTrxCD.Text), 1)
        Case "A", "B", "C"
            MsgBox "������ �� ���� �׸��Դϴ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
            Exit Sub
        Case Else
            ' no action
    End Select
    
    If Trim(txtControlTrxNM.Text) = "" Then
        MsgBox "�ð����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
        Exit Sub
    End If
    
    If sKaeyol = "" Then                '< 2007.12.18 : �迭�߰�
        MsgBox "��ȸ �� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
        Exit Sub
    End If
    
    If MsgBox("�����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� ����") = vbNo Then
        Exit Sub
    End If

    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  = '" & Trim(txtControlTrxCD.Text) & "'"
    sStr = sStr & "     AND KAEYOL = '" & sKaeyol & "'"     '< 2007.12.18 : �迭�߰�
    
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp =  Trim(txtControlTrxCD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        
        sGbnTrx = Left(Trim(txtControlTrxCD.Text), 1)
        
        sStr = ""
        sStr = sStr & "  SELECT TRXNM||'                              [T]'||TRXCD||'[T]'||TRX_CL AS TRX"
        sStr = sStr & "    FROM (SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND TRXCD  LIKE '" & sGbnTrx & "%'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"      '< 2007.12.18 : �迭�߰�
        sStr = sStr & "         UNION ALL"
        sStr = sStr & "         SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"      '< 2007.12.18 : �迭�߰�
        sStr = sStr & "            AND TRXCD  LIKE 'PB%'"
        sStr = sStr & "          )"
        sStr = sStr & "   ORDER BY TRXCD"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
            
    '    ' ACID
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    ' LSNTYPE
    '        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            cboTrx.Clear
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    If IsNull(.Fields("TRX")) = False Then
                        sTmp = Trim(.Fields("TRX"))
                            cboTrx.AddItem sTmp
                    End If
                    
                    .MoveNext
                Next nRec
            End If
        End With
    
        If cboTrx.ListCount > 0 Then cboTrx.ListIndex = 0
        
        fraTrx.Visible = False
        MsgBox "������ �ð������� �����Ͽ����ϴ�..", vbInformation + vbOKOnly, "������ �ð����� ����"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "������ �ð����� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ����"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "������ �ð����� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ����"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub

'>> ������ �ð����� ����
Private Sub cmdControlUpdateTrx_Click()
    Dim DBCmd       As ADODB.Command            '<< �л� �� ���� ����ϱ�
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nEXE        As Long
    Dim ni          As Long
    Dim nRec        As Long
    
    Dim sTrxCD      As String
    Dim nindex      As Integer
    Dim sGbnTrx     As String
    
    Dim sDiv()      As String
    
    On Error GoTo ErrStmt
    If Trim(txtControlTrxNM.Text) = "" Then
        MsgBox "�ð����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "������ �ð����� ���ŵ��"
        Exit Sub
    End If
    
    If sKaeyol = "" Then                '< 2007.12.18 : �迭�߰�
        MsgBox "��ȸ �� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
        Exit Sub
    End If
    
    
    If MsgBox("������ �½��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� ���ŵ��") = vbNo Then
        Exit Sub
    End If

    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    sStr = ""
    sStr = sStr & "  Update SDTRX01TB"
    sStr = sStr & "     SET TRXNM  = '" & Trim(txtControlTrxNM.Text) & "',"
    sStr = sStr & "         TRX_CL = " & lblControlTrxColor.BackColor
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  = '" & Trim(txtControlTrxCD.Text) & "'"
    sStr = sStr & "     AND KAEYOL = '" & sKaeyol & "'"     '< 2007.12.18 : �迭�߰�
    
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> �ð����� ����
'        sTmp = Trim(txtControlTrxNM.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXNM", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> color
'        nTmp = lblControlTrxColor.BackColor
'            Set DBParam = DBCmd.CreateParameter("TRX_CL", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp =  Trim(txtControlTrxCD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        
        sGbnTrx = Left(Trim(txtControlTrxCD.Text), 1)
        
        sStr = ""
        sStr = sStr & "  SELECT TRXNM||'                              [T]'||TRXCD||'[T]'||TRX_CL AS TRX"
        sStr = sStr & "    FROM (SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND TRXCD  LIKE '" & sGbnTrx & "%'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"                  '< 2007.12.18 : �迭�߰�
        sStr = sStr & "         UNION ALL"
        sStr = sStr & "         SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND TRXCD  LIKE 'PB%'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"                  '< 2007.12.18 : �迭�߰�
        sStr = sStr & "          )"
        sStr = sStr & "   ORDER BY TRXCD"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
            
    '    ' ACID
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    ' LSNTYPE
    '        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            cboTrx.Clear
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    If IsNull(.Fields("TRX")) = False Then
                        sTmp = Trim(.Fields("TRX"))
                        cboTrx.AddItem sTmp
                    End If
                    
                    .MoveNext
                Next nRec
            End If
        End With
    
        If cboTrx.ListCount > 0 Then cboTrx.ListIndex = 0
        
        fraTrx.Visible = False
        MsgBox "������ �ð����� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �ð����� ���ŵ��"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "������ �ð����� ���Ž� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���ŵ��"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "������ �ð����� ���Ž� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���ŵ��"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub

'>> ������ �ð����� ��� : ����� ����κи� PB XX �� �����ϴ� �ڵ�
Private Sub cmdControlInsertTrx_Click()
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nEXE        As Long
    Dim ni          As Long
    Dim nRec        As Long
    
    Dim sTrxCD      As String
    Dim nindex      As Integer
    Dim sGbnTrx     As String
    
    Dim sDiv()      As String
    
    On Error GoTo ErrStmt
    If Trim(txtControlTrxNM.Text) = "" Then
        MsgBox "�ð����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "������ �ð����� �űԵ��"
        Exit Sub
    End If
    
    If sKaeyol = "" Then                '< 2007.12.18 : �迭�߰�
        MsgBox "��ȸ �� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
        Exit Sub
    End If
    
    If MsgBox("���볻�� �߰��� �½��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� �űԵ��") = vbNo Then
        Exit Sub
    End If

    basDataBase.DBConn.BeginTrans

    sStr = ""
    sStr = sStr & "  SELECT 'PB'||TRIM(TO_CHAR(TO_NUMBER(MAX(SUBSTR(TRXCD,3,2))) + 1,'00')) AS TRXCD"
    sStr = sStr & "    FROM SDTRX01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  LIKE 'PB%' "
    sStr = sStr & "     AND KAEYOL = '" & sKaeyol & "'"     '< 2007.12.18 : �迭�߰�
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
                
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
            
    With DBRec
        If .RecordCount = 0 Then
            sTrxCD = "PB01"
        Else
            .MoveFirst
            If IsNull(.Fields("TRXCD")) = False Then
                sTrxCD = Trim(.Fields("TRXCD"))
            Else
                sTrxCD = "PB01"
            End If
        End If
    End With
    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX01TB (ACID, TRXCD, KAEYOL, TRXNM, TRX_CL) "           '< 2007.12.18 : �迭�߰�
    sStr = sStr & "         VALUES ("
    sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                 '" & sTrxCD & "',"
    sStr = sStr & "                 '" & sKaeyol & "',"                                     '< 2007.12.18 : �迭�߰�
    sStr = sStr & "                 '" & Trim(txtControlTrxNM.Text) & "',"
    sStr = sStr & "                 " & lblControlTrxColor.BackColor & ""
    sStr = sStr & "         )"
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = sTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �ð����� ����
'        sTmp = Trim(txtControlTrxNM.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXNM", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> color
'        nTmp = lblControlTrxColor.BackColor
'            Set DBParam = DBCmd.CreateParameter("TRX_CL", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        
        sGbnTrx = ""
        Me.Tag = "LOAD"
            nindex = cboTrx.ListIndex
            cboTrx.ListIndex = 0
            sDiv() = Split(cboTrx.Text, "[T]", -1, vbTextCompare)
        Me.Tag = ""
        
        If UBound(sDiv) <> 2 Then Exit Sub
        sGbnTrx = Left(Trim(sDiv(1)), 1)
        
        
        sStr = ""
        sStr = sStr & "  SELECT TRXNM||'                              [T]'||TRXCD||'[T]'||TRX_CL AS TRX"
        sStr = sStr & "    FROM (SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"              '< 2007.12.18 : �迭�߰�
        sStr = sStr & "            AND TRXCD  LIKE '" & sGbnTrx & "%'"
        sStr = sStr & "         UNION ALL"
        sStr = sStr & "         SELECT TRXNM, TRXCD, TRX_CL"
        sStr = sStr & "           FROM SDTRX01TB"
        sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "            AND TRXCD  LIKE 'PB%'"
        sStr = sStr & "            AND KAEYOL = '" & sKaeyol & "'"              '< 2007.12.18 : �迭�߰�
        sStr = sStr & "          )"
        sStr = sStr & "   ORDER BY TRXCD"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
            
    '    ' ACID
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    ' LSNTYPE
    '        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            cboTrx.Clear
            
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    If IsNull(.Fields("TRX")) = False Then
                        sTmp = Trim(.Fields("TRX"))
                            cboTrx.AddItem sTmp
                    End If
                    
                    .MoveNext
                Next nRec
            End If
        End With
    
        If cboTrx.ListCount > 0 Then cboTrx.ListIndex = 0
        
        fraTrx.Visible = False
        MsgBox "������ �ð����� �űԵ���Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �ð����� �űԵ��"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "������ �ð����� �űԵ�Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� �űԵ��"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "������ �ð����� �űԵ�Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� �űԵ��"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub





'>> ������ �ð�ǥ ���� ����
Private Sub SaveTrxColor(ByVal aTrxCD As String, ByVal aColor As Long)
    
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nEXE        As Long
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt

    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    sStr = ""
    sStr = sStr & " UPDATE SDTRX01TB "
    sStr = sStr & "    SET TRX_CL = " & CStr(aColor)
    sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"      '< 2007.12.18 : �迭�߰�
    sStr = sStr & "    AND TRXCD  = '" & aTrxCD & "'"

    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> color
'        nTmp = aColor
'            Set DBParam = DBCmd.CreateParameter("TRX_CL", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = aTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop



    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "������ ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� �����ϱ�"
        
        
        Me.Tag = "LOAD"
            
            sStr = ""
            sStr = sStr & "  SELECT TRXNM||'                              [T]'||TRXCD||'[T]'||TRX_CL AS TRX"
            sStr = sStr & "    FROM (SELECT TRXNM, TRXCD, TRX_CL"
            sStr = sStr & "           FROM SDTRX01TB"
            sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "            AND TRXCD  LIKE '" & Left(Trim(txtTrxCD.Text), 1) & "%'"
            sStr = sStr & "            AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"      '< 2007.12.18 : �迭�߰�
            sStr = sStr & "         UNION ALL"
            sStr = sStr & "         SELECT TRXNM, TRXCD, TRX_CL"
            sStr = sStr & "           FROM SDTRX01TB"
            sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "            AND TRXCD  LIKE 'PB%'"
            sStr = sStr & "            AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"      '< 2007.12.18 : �迭�߰�
            sStr = sStr & "          )"
            sStr = sStr & "   ORDER BY TRXCD"
            
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
                
        '    ' ACID
        '        sTmp = Trim(basModule.SchCD)
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '    ' LSNTYPE
        '        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            With DBRec
                cboTrx.Clear
                
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    For nRec = 1 To .RecordCount Step 1
                        If IsNull(.Fields("TRX")) = False Then
                            sTmp = Trim(.Fields("TRX"))
                                cboTrx.AddItem sTmp
                        End If
                        
                        .MoveNext
                    Next nRec
                End If
            End With
        
            If cboTrx.ListCount > 0 Then cboTrx.ListIndex = 0
            
            
        Me.Tag = ""
                
                
                
                
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "���� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �����ϱ�"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "���� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �����ϱ�"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub


Private Sub lblControlTrxColor_Click()
    
    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
    
        lblControlTrxColor.BackColor = .color
         
    End With
    
    Exit Sub
ErrStmt:

End Sub

Private Sub lblTrxColor_Click()

    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
    
        lblTrxColor.BackColor = .color
         
        Call SaveTrxColor(Trim(txtTrxCD.Text), .color)
         
    End With
    
    Exit Sub
ErrStmt:
    
End Sub

Private Sub cboTrx_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    Dim sDiv()      As String
    
    If Trim(cboTrx.Text) = "" Then Exit Sub
    
    sDiv = Split(cboTrx.Text, "[T]", -1, vbTextCompare)
    
    If UBound(sDiv) <> 2 Then Exit Sub
    
    txtTrxNM.Text = Trim(sDiv(0))
    txtTrxCD.Text = Trim(sDiv(1))
    lblTrxColor.BackColor = CLng(sDiv(2))
    
End Sub

Private Sub cboLsnType_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nCls        As Integer
    
    Dim sGbn        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT TRXNM||'                              [T]'||TRXCD||'[T]'||TRX_CL AS TRX"
    sStr = sStr & "    FROM (SELECT TRXNM, TRXCD, TRX_CL"
    sStr = sStr & "           FROM SDTRX01TB"
    sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND TRXCD  LIKE '" & Left(Trim(Right(cboLsnType, 30)), 1) & "%'"
    sStr = sStr & "            AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"      '< 2007.12.18 : �迭�߰�
    sStr = sStr & "         UNION ALL"
    sStr = sStr & "         SELECT TRXNM, TRXCD, TRX_CL"
    sStr = sStr & "           FROM SDTRX01TB"
    sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND TRXCD  LIKE 'PB%'"
    sStr = sStr & "            AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"      '< 2007.12.18 : �迭�߰�
    sStr = sStr & "          )"
    sStr = sStr & "   ORDER BY TRXCD"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' LSNTYPE
'        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        cboTrx.Clear
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                If IsNull(.Fields("TRX")) = False Then
                    sTmp = Trim(.Fields("TRX"))
                        cboTrx.AddItem sTmp
                End If
                
                .MoveNext
            Next nRec
        End If
    End With

    If cboTrx.ListCount > 0 Then cboTrx.ListIndex = 0

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð����� ����ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ����"
End Sub

Private Sub cmdControlTrx_Click()
    Dim sDiv()      As String
    
    fraTrx.Left = fraMain.Left + 200
    fraTrx.Top = fraMain.Top + cmdControlTrx.Top + cmdControlTrx.Height + 50
    fraTrx.ZOrder 0
    fraTrx.Visible = True
    
    txtControlTrxNM.SetFocus
    
    If Trim(cboTrx.Text) = "" Then Exit Sub
    
    sDiv = Split(cboTrx.Text, "[T]", -1, vbTextCompare)
    
    If UBound(sDiv) <> 2 Then Exit Sub
    
    txtControlTrxNM.Text = Trim(sDiv(0))
    txtControlTrxCD.Text = Trim(sDiv(1))
    lblControlTrxColor.BackColor = CLng(sDiv(2))
    
End Sub

Private Sub fraMain_Click()
    fraTrx.Visible = False
End Sub

Private Sub Frame2_Click()
    fraTrx.Visible = False
End Sub

Private Sub Frame3_Click()
    fraTrx.Visible = False
End Sub

Private Sub Frame4_Click()
    fraTrx.Visible = False
End Sub


Private Sub sprTrx_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim sDiv()      As String
    Dim sProc       As String
    Dim sTmp        As String
    Dim nWeekDay    As Integer
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sStr        As String
    Dim nEXE        As Long
    Dim bDeleteTrue As Boolean
    Dim sWeekday    As String
    
    If (Col = SpreadHeader) Or _
       (Row = SpreadHeader) Or _
       (Col < 1) Or _
       (Row < 1) Then
       
        Exit Sub
        
    End If
    
    fraTrx.Visible = False
    
    sProc = Trim(cmdTrxSel.Tag)         ' select : ���� / delete : ����
    sDiv = Split(cboTrx, "[T]", -1, vbTextCompare)
    
    If UBound(sDiv) <> 2 Then Exit Sub
    
    Select Case UCase(sProc)
        Case "SELECT"
            With sprTrx
            
                Select Case Col
                    Case 1
                        nWeekDay = 2
                    Case 2
                        nWeekDay = 3
                    Case 3
                        nWeekDay = 4
                    Case 4
                        nWeekDay = 5
                    Case 5
                        nWeekDay = 6
                    Case 6
                        nWeekDay = 7
                    Case 7
                        nWeekDay = 1
                End Select
            
                Select Case Find_Early_Save_Data(Row, nWeekDay)
                    Case "IN"
                        If Save_Setting_Data(Trim(txtTrxCD.Text), Row, nWeekDay) = True Then
                            .Row = Row
                            .Col = Col
                            If Len(Trim(.Text)) > 0 Then
                                sTmp = Trim(.Text)
                                sTmp = sTmp & vbCrLf & Trim(sDiv(0))
                            Else
                                sTmp = Trim(sDiv(0))
                            End If
                            
                                Call basFunction.Set_SprType_Text(sprTrx, "top", "left", basFunction.LenKor(sTmp), sTmp)
                                .Row2 = Row
                                .Col2 = Col
                                .BlockMode = True
                                    .BackColor = lblTrxColor.BackColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        End If
                    
                    
                    Case "NOT"
                        ' no action
                        
                End Select
                
            End With
            
        Case "DELETE"
            
            If chkAll.Value = 1 Then
                MsgBox "��ü ������ ���� üũ�� ���ְ�," & vbCrLf & _
                       "�����ϰ��� �ϴ� �� ���¸� ������ �ٽ� ��ȸ�Ͻʽÿ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
                Exit Sub
            End If
            
            
            
            On Error GoTo ErrStmt
            
            Select Case Col
                Case 1
                    nWeekDay = 2
                Case 2
                    nWeekDay = 3
                Case 3
                    nWeekDay = 4
                Case 4
                    nWeekDay = 5
                Case 5
                    nWeekDay = 6
                Case 6
                    nWeekDay = 7
                Case 7
                    nWeekDay = 1
            End Select
                        
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
            sStr = ""
            sStr = sStr & "  SELECT B.ACID, B.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, B.KAEYOL "                 '< 2007.12.18 : �迭�߰�
            sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
            sStr = sStr & "   WHERE A.ACID   = B.ACID "
            sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
            sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                                                    '< 2007.12.18 : �迭�߰�
            sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND B.TRXCD  LIKE '" & Left(Trim(Right(cboLsnType.Text, 30)), 1) & "%'"
            sStr = sStr & "     AND A.KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"                   '< 2007.12.18 : �迭�߰�
            sStr = sStr & "     AND B.LESSON = " & Trim(CStr(Row))
            sStr = sStr & "     AND B.WEEKS  = " & Trim(CStr(nWeekDay))
                    
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30

            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
                    
            '    ' ACID
            '        sTmp = Trim(basModule.SchCD)
            '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '    '>> lesson
            '        nTmp = row
            '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            '    '>> week
            '        nTmp = nweekday
            '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            
            bDeleteTrue = False     '<< �������ɿ���
            
            With DBRec
                
                Select Case .RecordCount
                    Case 0
                        
                        sStr = ""
                        sStr = sStr & "  SELECT B.ACID, B.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, B.KAEYOL "     '< 2007.12.18 : �迭�߰�
                        sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
                        sStr = sStr & "   WHERE A.ACID   = B.ACID "
                        sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
                        sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                                        '< 2007.12.18 : �迭�߰�
                        sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
                        sStr = sStr & "     AND B.TRXCD  LIKE 'PB%'"
                        sStr = sStr & "     AND A.KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"       '< 2007.12.18 : �迭�߰�
                        sStr = sStr & "     AND B.LESSON = " & Trim(CStr(Row))
                        sStr = sStr & "     AND B.WEEKS  = " & Trim(CStr(nWeekDay))
                                
                        Set DBCmd = New ADODB.Command
                        Set DBRec = New ADODB.Recordset
                        Set DBParam = New ADODB.Parameter
                        
                        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                        DBCmd.CommandText = sStr
                        DBCmd.CommandType = adCmdText
                        DBCmd.CommandTimeout = 30
            
                        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
                        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                            DBCmd.Parameters.Delete (0)
                        Next ni
                                
                        '    ' ACID
                        '        sTmp = Trim(basModule.SchCD)
                        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                        '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                        '    '>> lesson
                        '        nTmp = row
                        '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                        '    '>> week
                        '        nTmp = nweekday
                        '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
                        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
                        Do While DBRec.State And adStateExecuting
                            DoEvents
                        Loop
                        
                        With DBRec
                            
                            Select Case .RecordCount
                                Case 0
                                    MsgBox "������ �����Ͱ� �����ϴ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
                                    
                                Case Is = 1
                                    Select Case Trim(.Fields("WEEKS"))
                                        Case 2
                                            sWeekday = "��"
                                        Case 3
                                            sWeekday = "ȭ"
                                        Case 4
                                            sWeekday = "��"
                                        Case 5
                                            sWeekday = "��"
                                        Case 6
                                            sWeekday = "��"
                                        Case 7
                                            sWeekday = "��"
                                        Case 1
                                            sWeekday = "��"
                                    End Select
                                    
                                    If MsgBox("������ �����" & vbCrLf & _
                                              sWeekday & "���� - " & Trim(CStr(.Fields("LESSON"))) & "����" & vbCrLf & _
                                              Trim(.Fields("TRXNM")) & _
                                              "�� �½��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� ����") = vbYes Then
                                               
                                        bDeleteTrue = True
                                    End If
                                               
                                Case Else
                                    Select Case Trim(.Fields("WEEKS"))
                                        Case 2
                                            sWeekday = "��"
                                        Case 3
                                            sWeekday = "ȭ"
                                        Case 4
                                            sWeekday = "��"
                                        Case 5
                                            sWeekday = "��"
                                        Case 6
                                            sWeekday = "��"
                                        Case 7
                                            sWeekday = "��"
                                        Case 1
                                            sWeekday = "��"
                                    End Select
                                    
                                    If MsgBox("������ �����" & vbCrLf & _
                                              sWeekday & "���� - " & Trim(CStr(.Fields("LESSON"))) & "����" & vbCrLf & _
                                              Trim(.Fields("TRXNM")) & _
                                              "�� �½��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� ����") = vbYes Then
                                               
                                        bDeleteTrue = True
                                    End If
                                    
                            End Select
                        End With
                        
                        
                    Case Is = 1
                        Select Case Trim(.Fields("WEEKS"))
                            Case 2
                                sWeekday = "��"
                            Case 3
                                sWeekday = "ȭ"
                            Case 4
                                sWeekday = "��"
                            Case 5
                                sWeekday = "��"
                            Case 6
                                sWeekday = "��"
                            Case 7
                                sWeekday = "��"
                            Case 1
                                sWeekday = "��"
                        End Select
                    
                        If MsgBox("������ �����" & vbCrLf & _
                                  sWeekday & "���� - " & Trim(CStr(.Fields("LESSON"))) & "����" & vbCrLf & _
                                  Trim(.Fields("TRXNM")) & _
                                  "�� �½��ϱ�?", vbQuestion + vbYesNo, "������ �ð����� ����") = vbYes Then
                                               
                            bDeleteTrue = True
                        End If
                            
                    Case Else
                        MsgBox "������ �����Ͱ� ��Ȯ���� �ʽ��ϴ�." & _
                               "�����ڿ��� ���ǹٶ��ϴ�.", vbExclamation + vbOKOnly, "������ �ð����� ����"
                        
                End Select
                
                If bDeleteTrue = True Then
            
                    With DBRec
                        .MoveFirst
                            
                            basDataBase.DBConn.BeginTrans
        
                            sStr = ""
                            sStr = sStr & "  DELETE "
                            sStr = sStr & "    FROM SDTRX11TB "
                            sStr = sStr & "   WHERE ACID   = '" & Trim(.Fields("ACID")) & "'"
                            sStr = sStr & "     AND TRXCD  = '" & Trim(.Fields("TRXCD")) & "'"
                            sStr = sStr & "     AND KAEYOL = '" & Trim(.Fields("KAEYOL")) & "'"         '< 2007.12.18 : �迭�߰�
                            sStr = sStr & "     AND LESSON = " & Trim(.Fields("LESSON"))
                            sStr = sStr & "     AND WEEKS  = " & Trim(.Fields("WEEKS"))
                            
                            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
                            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                                DBCmd.Parameters.Delete (0)
                            Next ni
                            
                            
                        '    '>> �п�
                        '        sTmp = Trim(.Fields("ACID"))
                        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                        '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                        '    '>> ������ �ð�ǥ ����
                        '        sTmp = Trim(.Fields("TRXCD"))
                        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                        '            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                        '    '>> LESSON
                        '        nTmp = Trim(.Fields("LESSON"))
                        '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                        '    '>> WEEKS
                        '        nTmp = Trim(.Fields("WEEKS"))
                        '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                        
                            DBCmd.CommandText = sStr
                            DBCmd.CommandType = adCmdText
                            DBCmd.CommandTimeout = 30
                        
                            nEXE = 0
                            DBCmd.Execute nEXE, , -1
                        
                            Do While basDataBase.DBConn.State And adStateExecuting
                                DoEvents
                            Loop
                        
                            If nEXE = 1 Then
                                basDataBase.DBConn.CommitTrans
                                
                                sprTrx.Row = Row
                                sprTrx.Col = Col
                                    sTmp = ""
                                    Call basFunction.Set_SprType_Text(sprTrx, "top", "left", 1, sTmp)
                                    sprTrx.BackColor = basModule.WhiteColor
                                    sprTrx.BackColorStyle = BackColorStyleUnderGrid
                            Else
                                basDataBase.DBConn.RollbackTrans
                            End If
                    End With
                End If
                
                'chkAll.Value = 1
                Call cmdFindMtx_Click
                
            End With
    End Select
    
    
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    
    MsgBox "ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð����� ��� �� ����"
    
End Sub


Private Function Delete_Setting_Data(ByVal aTrxCD As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As Boolean
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    Dim ni          As Integer
    Dim nEXE        As Integer
    
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX11TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  = '" & aTrxCD & "'"
    sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30))       '< 2007.12.18 : �迭�߰�
    sStr = sStr & "     AND LESSON = " & Trim(CStr(aLesson))
    sStr = sStr & "     AND WEEKS  = " & Trim(CStr(aWeek))
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = aTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> LESSON
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> WEEKS
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
    Else
        basDataBase.DBConn.RollbackTrans
        bRet = False
    End If
        
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Delete_Setting_Data = bRet
    
    Exit Function
ErrStmt:

    basDataBase.DBConn.RollbackTrans

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð����� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
    
    Delete_Setting_Data = bRet
    
End Function


'>> �����.
Private Function Save_Setting_Data(ByVal aTrxCD As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As Boolean
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    Dim ni          As Integer
    Dim nEXE        As Integer
    
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX11TB (ACID, TRXCD, KAEYOL, LESSON, WEEKS)"        '< 2007.12.18 : �迭�߰�
    sStr = sStr & "  VALUES("
    sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                 '" & aTrxCD & "',"
    sStr = sStr & "                 '" & Trim(Right(cboKaeyol.Text, 30)) & "',"         '< 2007.12.18 : �迭�߰�
    sStr = sStr & "                  " & Trim(CStr(aLesson)) & ","
    sStr = sStr & "                  " & Trim(CStr(aWeek))
    sStr = sStr & "         )"
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni

'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = aTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> LESSON
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> WEEKS
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nEXE = 0
    DBCmd.Execute nEXE, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nEXE = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
    Else
        basDataBase.DBConn.RollbackTrans
        bRet = False
    End If
        
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Save_Setting_Data = bRet
    
    Exit Function
ErrStmt:
    
    If Err.Number = -2147217900 Then
        'MsgBox "update"
        On Error GoTo 0
        On Error GoTo ErrUpdate
        
        sStr = ""
        sStr = sStr & "  DELETE "
        sStr = sStr & "    FROM SDTRX11TB "
        sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TRXCD  = '" & aTrxCD & "'"
        sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"         '< 2007.12.18 : �迭�߰�
        sStr = sStr & "     AND LESSON = " & Trim(CStr(aLesson))
        sStr = sStr & "     AND WEEKS  = " & Trim(CStr(aWeek))
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> ������ �ð�ǥ ����
    '        sTmp = aTrxCD
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> LESSON
    '        nTmp = aLesson
    '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '    '>> WEEKS
    '        nTmp = aWeek
    '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
    
        nEXE = 0
        DBCmd.Execute nEXE, , -1
    
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
    
        If nEXE = 1 Then
            
            sStr = ""
            sStr = sStr & "  INSERT INTO SDTRX11TB (ACID, TRXCD, KAEYOL, LESSON, WEEKS)"    '< 2007.12.18 : �迭�߰�
            sStr = sStr & "  VALUES("
            sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
            sStr = sStr & "                 '" & aTrxCD & "',"
            sStr = sStr & "                 '" & Trim(Right(cboKaeyol.Text, 30)) & "',"     '< 2007.12.18 : �迭�߰�
            sStr = sStr & "                  " & Trim(CStr(aLesson)) & ","
            sStr = sStr & "                  " & Trim(CStr(aWeek))
            sStr = sStr & "         )"
            
            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
        
        '    '>> �п�
        '        sTmp = Trim(basModule.SchCD)
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        '    '>> ������ �ð�ǥ ����
        '        sTmp = aTrxCD
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        '    '>> LESSON
        '        nTmp = aLesson
        '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '    '>> WEEKS
        '        nTmp = aWeek
        '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
        
            nEXE = 0
            DBCmd.Execute nEXE, , -1
        
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
        
            If nEXE = 1 Then
                basDataBase.DBConn.CommitTrans
                bRet = True
            Else
                basDataBase.DBConn.RollbackTrans
                bRet = False
            End If
        Else
            basDataBase.DBConn.RollbackTrans
        
            Set DBCmd = Nothing
            Set DBRec = Nothing
            
            On Error GoTo 0
            MsgBox "������ �ð����� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
        End If
        
    Else
        basDataBase.DBConn.RollbackTrans
    
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        On Error GoTo 0
        MsgBox "������ �ð����� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
        
    End If
    
    Save_Setting_Data = bRet
    
    Exit Function
ErrUpdate:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð����� ��Ͽ����� �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
    
    Save_Setting_Data = bRet
        
End Function



'>> ���� ��ϵ� ������ �ִ��� Ȯ����.
Private Function Find_Early_Save_Data(ByVal aLesson As Integer, ByVal aWeek As Integer) As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sStr        As String
    Dim sRet        As String
    Dim sTmp        As String
    Dim sWeekday    As String
    
    sStr = ""
    sStr = sStr & "  SELECT B.ACID, B.TRXCD, A.TRXNM, A.KAEYOL"                     '< 2007.12.18 : �迭�߰�
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   WHERE A.ACID   = B.ACID "
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"   '< 2007.12.18 : �迭�߰�
    sStr = sStr & "     AND B.LESSON = " & Trim(CStr(aLesson))
    sStr = sStr & "     AND B.WEEKS  = " & Trim(CStr(aWeek))
            
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> lesson
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> week
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sRet = "NOT"
    With DBRec
        
        Select Case .RecordCount
            Case 0
                sRet = "IN"
                
            Case Is > 0
                .MoveFirst
                    
                sTmp = ""
                For nRec = 1 To .RecordCount Step 1
                    If IsNull(.Fields("TRXNM")) = False Then
                        If nRec > 1 Then sTmp = sTmp & vbCrLf
                        sTmp = sTmp & Trim(.Fields("TRXNM"))
                    End If
                    
                    .MoveNext
                Next nRec
                
                Select Case aWeek
                    Case 2
                        sWeekday = "��"
                    Case 3
                        sWeekday = "ȭ"
                    Case 4
                        sWeekday = "��"
                    Case 5
                        sWeekday = "��"
                    Case 6
                        sWeekday = "��"
                    Case 7
                        sWeekday = "��"
                    Case 1
                        sWeekday = "��"
                End Select
                
                sTmp = sWeekday & "���� - " & Trim(CStr(aLesson)) & "����" & vbCrLf & vbCrLf & sTmp
                If MsgBox(sTmp & vbCrLf & "������ �ֽ��ϴ�. �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "���� ��ϳ��� ��ȸ") = vbYes Then
                    sRet = "IN"
                Else
                    sRet = "NOT"
                End If
                
        End Select
    End With
    
    Find_Early_Save_Data = sRet

End Function


