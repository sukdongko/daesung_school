VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMR011 
   Caption         =   "�ð�ǥ ����� >> ���� �� ���纰 ����ֱ�"
   ClientHeight    =   9975
   ClientLeft      =   7500
   ClientTop       =   3810
   ClientWidth     =   7905
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   7905
   Begin VB.Frame Frame1 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame1"
      Height          =   2385
      Left            =   60
      TabIndex        =   12
      Top             =   30
      Width           =   7725
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   2325
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   7665
         Begin VB.CommandButton cmdNewTeacher 
            Caption         =   "�� ��"
            Height          =   400
            Left            =   990
            TabIndex        =   0
            Top             =   180
            Width           =   1000
         End
         Begin VB.CommandButton cmdDeleteTeacher 
            Caption         =   "�� ��"
            Height          =   400
            Left            =   4530
            TabIndex        =   3
            Top             =   180
            Width           =   1000
         End
         Begin VB.CommandButton cmdSaveTeacher 
            Caption         =   "�� ��(&S)"
            Height          =   400
            Left            =   3330
            TabIndex        =   2
            Top             =   180
            Width           =   1000
         End
         Begin VB.CommandButton cmdFindTeacher 
            Caption         =   "�� ȸ"
            Height          =   400
            Left            =   2160
            TabIndex        =   1
            Top             =   180
            Width           =   1000
         End
         Begin VB.TextBox txtSubjNM 
            Height          =   300
            IMEMode         =   10  '�ѱ� 
            Left            =   1830
            TabIndex        =   7
            Text            =   "txtSubjNM"
            Top             =   1110
            Width           =   1455
         End
         Begin VB.ComboBox cboSubjGbn 
            Height          =   300
            Left            =   1200
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   8
            Top             =   1485
            Width           =   1455
         End
         Begin VB.TextBox txtTcrNM 
            Height          =   300
            IMEMode         =   10  '�ѱ� 
            Left            =   1830
            TabIndex        =   5
            Text            =   "txtTcrNM"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox cboTcrGbn 
            Height          =   300
            Left            =   1200
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   9
            Top             =   1875
            Width           =   1455
         End
         Begin EditLib.fpMask fpTcrCD 
            Height          =   300
            Left            =   1200
            TabIndex        =   4
            Top             =   720
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "999"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   0   'False
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            AutoTab         =   0   'False
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpMask fpSubjCD 
            Height          =   300
            Left            =   1200
            TabIndex        =   6
            Top             =   1110
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
            BorderColor     =   -2147483642
            BorderWidth     =   1
            ButtonDisable   =   0   'False
            ButtonHide      =   0   'False
            ButtonIncrement =   1
            ButtonMin       =   0
            ButtonMax       =   100
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            AllowOverflow   =   0   'False
            BestFit         =   0   'False
            ClipMode        =   0
            DataFormatEx    =   0
            Mask            =   "99"
            PromptChar      =   "_"
            PromptInclude   =   0   'False
            RequireFill     =   0   'False
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            AutoTab         =   0   'False
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   5730
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   $"TMR011.frx":0000
            Height          =   360
            Left            =   3390
            TabIndex        =   19
            Top             =   1065
            Width           =   3765
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "���纰 ������ ��ȣ�� ����Ͻʽÿ�."
            Height          =   210
            Left            =   3390
            TabIndex        =   18
            Top             =   765
            Width           =   3765
         End
         Begin VB.Label Label2 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����"
            Height          =   210
            Left            =   150
            TabIndex        =   17
            Top             =   1155
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���񱸺�"
            Height          =   210
            Left            =   150
            TabIndex        =   16
            Top             =   1530
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����"
            Height          =   210
            Left            =   150
            TabIndex        =   15
            Top             =   765
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���籸��"
            Height          =   210
            Left            =   150
            TabIndex        =   14
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblTcrColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"TMR011.frx":003F
            Height          =   615
            Left            =   3540
            TabIndex        =   10
            Top             =   1590
            Width           =   765
         End
      End
   End
   Begin FPSpread.vaSpread sprTcr 
      Height          =   7095
      Left            =   30
      TabIndex        =   11
      Top             =   2460
      Width           =   7755
      _Version        =   393216
      _ExtentX        =   13679
      _ExtentY        =   12515
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
      MaxCols         =   10
      SpreadDesigner  =   "TMR011.frx":0055
   End
End
Attribute VB_Name = "TMR011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM011
'   �� ��  �� �� : ���� �� ���纰 ����ֱ�
'
'   ��   ��   �� : 2007/12/27
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit



Private Sub Form_Load()
    Me.Move 0, 0, 8000, 9980
    
    With sprTcr
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
    End With
    
    With cboSubjGbn
        .Clear
        .AddItem "���" & Space(50) & "10"
        .AddItem "����" & Space(50) & "20"
        .AddItem "�ܱ���" & Space(50) & "30"
        .AddItem "��Ž" & Space(50) & "40"
        .AddItem "��Ž" & Space(50) & "50"
        
        .ListIndex = 0
    End With
    
    With cboTcrGbn
        .Clear
        .AddItem "����" & Space(50) & "99"
        .AddItem "����" & Space(50) & "10"
        .AddItem "���� �Ⱝ" & Space(50) & "20"
        .AddItem "���� �Ⱝ" & Space(50) & "30"
        
        .ListIndex = 0
    End With
    
    Me.Tag = "LOAD"
        Call initData
    
    Me.Tag = ""

End Sub

Private Sub cmdNewTeacher_Click()
    Call initData
    
End Sub

Private Sub initData()
    fpTcrCD.Text = ""
    txtTcrNM.Text = ""
    
    fpSubjCD.Text = ""
    txtSubjNM.Text = ""
    
    lblTcrColor.BackColor = &HFFFFFF
End Sub

Private Sub lblTcrColor_Click()
    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
        lblTcrColor.BackColor = .color
    End With
    
    Exit Sub
ErrStmt:

End Sub


'>> ���� �� ���纰 ������ȸ
Private Sub cmdFindTeacher_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nColor      As Long
    
    sprTcr.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
    sStr = sStr & "    From SDTCR01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(txtTcrNM.Text) > " " Then
        sStr = sStr & " AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "'"
    End If
    If Trim(fpTcrCD.UnFmtText) > " " Then
        sStr = sStr & " AND TCRCD  LIKE '" & Trim(fpTcrCD.UnFmtText) & "'"
    End If
    If Trim(txtSubjNM.Text) > " " Then
        sStr = sStr & " AND SUBJNM LIKE '" & Trim(txtSubjNM.Text) & "'"
    End If
    If Trim(fpSubjCD.UnFmtText) > " " Then
        sStr = sStr & " AND SUBJCD LIKE '" & Trim(fpSubjCD.UnFmtText) & "'"
    End If
    sStr = sStr & "   ORDER BY ACID, TCRCD, SUBJCD "
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows:        sprTcr.RowHeight(sprTcr.Row) = 16
                
                sprTcr.Col = 1
                    sTmp = " ":  If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":  If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":  If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":  If IsNull(.Fields("SUBJGBN")) = False Then sTmp = Trim(.Fields("SUBJGBN"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                For ni = 0 To cboSubjGbn.ListCount - 1 Step 1
                    cboSubjGbn.ListIndex = ni
                    If InStr(1, Trim(cboSubjGbn.Text), sTmp, vbTextCompare) > 0 Then
                        sprTcr.Col = sprTcr.Col + 1
                            sTmp = Trim(Mid(cboSubjGbn.Text, 1, 40))
                            Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                        Exit For
                    End If
                Next ni
                
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TCRGBN")) = False Then sTmp = Trim(.Fields("TCRGBN"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                For ni = 0 To cboTcrGbn.ListCount - 1 Step 1
                    cboTcrGbn.ListIndex = ni
                        
                    If InStr(1, Trim(cboTcrGbn.Text), sTmp, vbTextCompare) > 0 Then
                        sprTcr.Col = sprTcr.Col + 1
                            sTmp = Trim(Mid(cboTcrGbn.Text, 1, 40))
                            If StrComp("����", sTmp, vbTextCompare) = 0 Then
                                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", 1, "")
                            Else
                                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                            End If
                        Exit For
                    End If
                Next ni
                
                sprTcr.Col = sprTcr.Col + 1
                    nColor = 0
                        If IsNumeric(.Fields("TCR_CL")) = True Then nColor = CLng(.Fields("TCR_CL"))
                        sprTcr.Row2 = sprTcr.Row
                        sprTcr.Col2 = sprTcr.Col
                        sprTcr.BlockMode = True
                            sprTcr.BackColor = nColor
                            sprTcr.BackColorStyle = BackColorStyleUnderGrid
                        sprTcr.BlockMode = False
                
                sprTcr.Col = sprTcr.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprTcr)
                
                .MoveNext
            Next nRec
        End If
    End With
    
    cboSubjGbn.ListIndex = 0
    cboTcrGbn.ListIndex = 0

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� �� ���纰 ����ֱ�"
    
    fpTcrCD.SetFocus
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� �� ���纰 ����ֱ� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �� ���纰 ����ֱ�"
End Sub


'>> ���� �� ���纰 ���񳻿� ���
Private Sub cmdSaveTeacher_Click()
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim sComp       As String
    Dim nExe        As Long
    
    Dim sSaveGbn    As String
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim ni          As Long
    Dim nRow        As Long
    Dim nColor      As Long
    
    If Trim(fpTcrCD.UnFmtText) = "" Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 3�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    If Len(fpTcrCD.UnFmtText) <> 3 Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 3�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    
    If Trim(fpSubjCD.UnFmtText) = "" Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 2�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    If Len(fpSubjCD.UnFmtText) <> 2 Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 2�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    
    If Trim(txtTcrNM.Text) = "" Then
        MsgBox "������� �����ϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    If Trim(txtSubjNM.Text) = "" Then
        MsgBox "������� �����ϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 ����ֱ�"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    nExe = 0
    
    sStr = ""
    sStr = sStr & "  SELECT TCRCD, SUBJCD"
    sStr = sStr & "    FROM SDTCR01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD  = '" & Trim(fpTcrCD.UnFmtText) & "'"
    sStr = sStr & "     AND SUBJCD = '" & Trim(fpSubjCD.UnFmtText) & "'"
            
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30


    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    
    With DBRec
        If .RecordCount = 0 Then
            sTcrCD = Trim(fpTcrCD.UnFmtText)
            sSubjCD = Trim(fpSubjCD.UnFmtText)
            
            sSaveGbn = "INSERT"
            
        ElseIf .RecordCount > 0 Then
            .MoveFirst
            
            sTcrCD = "":        If IsNull(.Fields("TCRCD")) = False Then sTcrCD = Trim(.Fields("TCRCD"))
            sSubjCD = "":       If IsNull(.Fields("SUBJCD")) = False Then sSubjCD = Trim(.Fields("SUBJCD"))
            
            sSaveGbn = "UPDATE"
            
        End If
    End With
                
    Set DBRec = Nothing
    
    If sSaveGbn = "INSERT" Then
        '<< INSERT >>
        sStr = ""
        sStr = sStr & "  INSERT INTO SDTCR01TB ( ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL ) "
        sStr = sStr & "  VALUES ( "
        sStr = sStr & "          '" & Trim(basModule.SchCD) & "',"
        sStr = sStr & "          '" & sTcrCD & "',"
        sStr = sStr & "          '" & sSubjCD & "',"
        sStr = sStr & "          '" & Trim(Right(cboSubjGbn.Text, 30)) & "',"
        sStr = sStr & "          '" & Trim(Right(cboTcrGbn.Text, 30)) & "',"
        sStr = sStr & "          '" & Trim(txtTcrNM.Text) & "',"
        sStr = sStr & "          '" & Trim(txtSubjNM.Text) & "',"
        sStr = sStr & "          " & Trim(CStr(lblTcrColor.BackColor))
        sStr = sStr & "  ) "
            
    ElseIf sSaveGbn = "UPDATE" Then
        '<< UPDATE >>
        sStr = ""
        sStr = sStr & "  UPDATE SDTCR01TB"
        sStr = sStr & "     SET SUBJGBN = '" & Trim(Right(cboSubjGbn.Text, 30)) & "',"
        sStr = sStr & "         TCRGBN  = '" & Trim(Right(cboTcrGbn.Text, 30)) & "',"
        sStr = sStr & "         TCRNM   = '" & Trim(txtTcrNM.Text) & "',"
        sStr = sStr & "         SUBJNM  = '" & Trim(txtSubjNM.Text) & "',"
        sStr = sStr & "         TCR_CL  = " & Trim(CStr(lblTcrColor.BackColor))
        sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TCRCD   = '" & sTcrCD & "'"
        sStr = sStr & "     AND SUBJCD  = '" & sSubjCD & "'"
        
        If MsgBox("�̹� ��ϵ� ������ �����Ƿ� ����ó���մϴ�." & vbCrLf & _
                   "����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            basDataBase.DBConn.RollbackTrans
    
            Set DBCmd = Nothing
            Set DBParam = Nothing
        End If
    End If
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
            
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        
        With sprTcr
        Select Case sSaveGbn
            Case "INSERT"
                .MaxRows = .MaxRows + 1
                .InsertRows 1, 1
                .Row = 1:           .RowHeight(.Row) = 16
            Case "UPDATE"
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = 1:   sTmp = Trim(.Text)              '< �����ڵ�
                    .Col = 3:   sTmp = sTmp & Trim(.Text)       '< �����ڵ�
                    
                    sComp = sTcrCD & sSubjCD
                    
                    If StrComp(sComp, sTmp, vbTextCompare) = 0 Then
                        .Row = nRow
                        Exit For
                    End If
                Next nRow
        End Select
        End With
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    '## ������ spread��...
    With sprTcr
    '> row �� ������ ���ǵ�.
        .Col = 1
            sTmp = sTcrCD
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
        .Col = .Col + 1
            sTmp = Trim(txtTcrNM.Text)
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
        
        .Col = .Col + 1
            sTmp = sSubjCD
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
        .Col = .Col + 1
            sTmp = Trim(txtSubjNM.Text)
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
        
        .Col = .Col + 1
            sTmp = Trim(Right(cboSubjGbn.Text, 30))
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
            '>> ���񱸺и�
            If InStr(1, Trim(cboSubjGbn), sTmp, vbTextCompare) > 0 Then
                .Col = .Col + 1
                    sTmp = Trim(Mid(cboSubjGbn.Text, 1, 40))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
            End If
        
        
        .Col = .Col + 1
            sTmp = Trim(Right(cboTcrGbn.Text, 30))
                Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
            '>> ���籸�и�
            If InStr(1, Trim(cboTcrGbn.Text), sTmp, vbTextCompare) > 0 Then
                .Col = .Col + 1
                    sTmp = Trim(Mid(cboTcrGbn.Text, 1, 40))
                    If StrComp("����", sTmp, vbTextCompare) = 0 Then
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", 1, "")
                    Else
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                    End If
            End If
        
        
        .Col = .Col + 1
            nColor = 0
            nColor = lblTcrColor.BackColor
            .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = nColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
        
        .Col = .Col + 1
            Call basFunction.Set_SprType_ChkBox(sprTcr)
            
        If .MaxRows > 0 Then
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End If
    End With
    
    Call initData
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� �� ���纰 ���񳻿� �ֱ�"
    
    fpTcrCD.SetFocus
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
End Sub

'>> ���� �� ���񳻿� ����
Private Sub cmdDeleteTeacher_Click()
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim sComp       As String
    Dim nExe        As Long
    
    Dim ni          As Long
    Dim nRow        As Long

    If Trim(fpTcrCD.UnFmtText) = "" Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 3�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    If Len(fpTcrCD.UnFmtText) <> 3 Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 3�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    
    If Trim(fpSubjCD.UnFmtText) = "" Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 2�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    If Len(fpSubjCD.UnFmtText) <> 2 Then
        MsgBox "�����ڵ带 ����Ͻʽÿ�." & vbCrLf & _
               "�����ڵ�� ���ڷ� 2�ڸ� �Դϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    
    If Trim(txtTcrNM.Text) = "" Then
        MsgBox "������� �����ϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    If Trim(txtSubjNM.Text) = "" Then
        MsgBox "������� �����ϴ�.", vbExclamation + vbOKOnly, "���� �� ���纰 �������"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    If MsgBox("����ó���Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "���� �� ���纰 �������") = vbNo Then
        Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    nExe = 0
    
    '<< DELETE >>
    sStr = ""
    sStr = sStr & "  DELETE SDTCR01TB"
    sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD   = '" & Trim(fpTcrCD.UnFmtText) & "'"
    sStr = sStr & "     AND SUBJCD  = '" & Trim(fpSubjCD.UnFmtText) & "'"
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
            
    If nExe = 1 Then
        
        On Error GoTo 0
        On Error Resume Next
        
        '<< DELETE >>
        sStr = ""
        sStr = sStr & "  DELETE SDTCR11TB"
        sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TCRCD   = '" & Trim(fpTcrCD.UnFmtText) & "'"
        sStr = sStr & "     AND SUBJCD  = '" & Trim(fpSubjCD.UnFmtText) & "'"
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
                        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        '<< DELETE >>
        sStr = ""
        sStr = sStr & "  DELETE SDTCR15TB"
        sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TCRCD   = '" & Trim(fpTcrCD.UnFmtText) & "'"
        sStr = sStr & "     AND SUBJCD  = '" & Trim(fpSubjCD.UnFmtText) & "'"
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
                        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
            
        '<< DELETE >>
        sStr = ""
        sStr = sStr & "  DELETE SDTRX50TB"
        sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TCRCD   = '" & Trim(fpTcrCD.UnFmtText) & "'"
        sStr = sStr & "     AND SUBJCD  = '" & Trim(fpSubjCD.UnFmtText) & "'"
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
                        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        basDataBase.DBConn.CommitTrans
        On Error GoTo 0
        
        
        With sprTcr
    
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 1:   sTmp = Trim(.Text)              '< �����ڵ�
            .Col = 3:   sTmp = sTmp & Trim(.Text)       '< �����ڵ�
            
            sComp = Trim(fpTcrCD.UnFmtText) & Trim(fpSubjCD.UnFmtText)
            
            If StrComp(sComp, sTmp, vbTextCompare) = 0 Then
                .Row = nRow
                .DeleteRows .Row, 1
                .MaxRows = .MaxRows - 1
            End If
        Next nRow
        End With
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    'Call initData
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "�����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� �� ���纰 �������"
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub


Private Sub sprTcr_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprTcr
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 8
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 10:          .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:    .Value = 0                      '< ���þ���.
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = 8
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 10:          .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        '>> ������ �����ֱ�
        
        .Row = Row
        .Col = 1:       fpTcrCD.Text = Trim(.Text)
        .Col = 2:       txtTcrNM.Text = Trim(.Text)
        .Col = 3:       fpSubjCD.Text = Trim(.Text)
        .Col = 4:       txtSubjNM.Text = Trim(.Text)
        
        For ni = 0 To cboSubjGbn.ListCount - 1 Step 1
            cboSubjGbn.ListIndex = ni
            
            .Col = 5
            If StrComp(Trim(Right(cboSubjGbn.Text, 30)), Trim(.Text), vbTextCompare) = 0 Then
                Exit For
            End If
        Next ni
        
        For ni = 0 To cboTcrGbn.ListCount - 1 Step 1
            cboTcrGbn.ListIndex = ni
            
            .Col = 7
            If StrComp(Trim(Right(cboTcrGbn.Text, 30)), Trim(.Text), vbTextCompare) = 0 Then
                Exit For
            End If
        Next ni
        
        .Col = 9
            lblTcrColor.BackColor = .BackColor
        .Col = .MaxCols
            .Value = 1                                      '< ����
        
    End With
End Sub

































