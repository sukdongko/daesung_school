VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LSN001 
   Caption         =   "�ð�ǥ ����� >> �� ���� ���"
   ClientHeight    =   9660
   ClientLeft      =   13890
   ClientTop       =   2115
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   8685
   Begin VB.Frame fraUpdate 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Height          =   5115
      Left            =   510
      TabIndex        =   2
      Top             =   3570
      Width           =   4845
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   5055
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   4785
         Begin VB.TextBox txtUpBase_Class 
            Height          =   360
            Left            =   1200
            TabIndex        =   16
            Text            =   "txtUpBase_Class"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtUpDamim 
            Height          =   360
            Left            =   1200
            TabIndex        =   15
            Text            =   "txtUpDamim"
            Top             =   2070
            Width           =   1455
         End
         Begin VB.TextBox txtUpLsnCDNM 
            Height          =   360
            Left            =   1200
            TabIndex        =   13
            Text            =   "txtUpLsnCDNM"
            Top             =   1230
            Width           =   1455
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "�� �����ϱ�"
            Height          =   500
            Left            =   2550
            TabIndex        =   20
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "�� �����ϱ�"
            Height          =   500
            Left            =   480
            TabIndex        =   19
            Top             =   4260
            Width           =   1665
         End
         Begin VB.TextBox txtUpLsnCD 
            Enabled         =   0   'False
            Height          =   360
            Left            =   1200
            TabIndex        =   11
            Text            =   "txtUpLsnCD"
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox cboUpKaeyol 
            Height          =   300
            Left            =   1200
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   14
            Top             =   1650
            Width           =   1485
         End
         Begin VB.TextBox txtUpLsnNM 
            Height          =   360
            Left            =   1200
            TabIndex        =   12
            Text            =   "txtUpLsnNM"
            Top             =   810
            Width           =   1455
         End
         Begin EditLib.fpLongInteger fpUpLsnCapa 
            Height          =   360
            Left            =   1200
            TabIndex        =   17
            Top             =   2940
            Width           =   1095
            _Version        =   196608
            _ExtentX        =   1931
            _ExtentY        =   635
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   1
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
            Text            =   "0"
            MaxValue        =   "99999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label13 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ǽ�"
            Height          =   210
            Left            =   90
            TabIndex        =   37
            Top             =   2610
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   90
            TabIndex        =   36
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label12 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   90
            TabIndex        =   35
            Top             =   3510
            Width           =   975
         End
         Begin VB.Label lblUpLsnColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"LSN001.frx":0000
            Height          =   675
            Left            =   1200
            TabIndex        =   18
            Top             =   3420
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "�ݱ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00CB5C56&
            Height          =   375
            Left            =   3900
            TabIndex        =   33
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label10 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ڵ��Ī"
            Height          =   210
            Left            =   90
            TabIndex        =   32
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ڵ�"
            Height          =   210
            Left            =   90
            TabIndex        =   25
            Top             =   450
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� ����                      ��"
            Height          =   210
            Left            =   -810
            TabIndex        =   24
            Top             =   3030
            Width           =   3375
         End
         Begin VB.Label Label6 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   90
            TabIndex        =   23
            Top             =   1725
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���̸�"
            Height          =   210
            Left            =   90
            TabIndex        =   22
            Top             =   900
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame3"
      Height          =   9525
      Left            =   60
      TabIndex        =   26
      Top             =   30
      Width           =   8505
      Begin VB.Frame Frame1 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   9465
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   8445
         Begin VB.TextBox txtBase_Class 
            Height          =   360
            Left            =   1380
            TabIndex        =   7
            Text            =   "txtBase_Class"
            Top             =   2070
            Width           =   1455
         End
         Begin VB.TextBox txtDamim 
            Height          =   360
            Left            =   4380
            TabIndex        =   6
            Text            =   "txtDamim"
            Top             =   1620
            Width           =   1455
         End
         Begin VB.TextBox txtLsnCDNM 
            Height          =   360
            Left            =   4380
            TabIndex        =   4
            Text            =   "txtLsnCDNM"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "�� ����ϱ� (&S)"
            Height          =   500
            Left            =   1350
            TabIndex        =   0
            Top             =   330
            Width           =   1665
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�� ��ȸ�ϱ� (&F)"
            Height          =   500
            Left            =   3300
            TabIndex        =   1
            Top             =   330
            Width           =   1665
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   1380
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   5
            Top             =   1680
            Width           =   1485
         End
         Begin VB.TextBox txtLsnNM 
            Height          =   360
            Left            =   1380
            TabIndex        =   3
            Text            =   "txtLsnNM"
            Top             =   1230
            Width           =   1455
         End
         Begin FPSpread.vaSpread sprDisp 
            Height          =   6735
            Left            =   90
            TabIndex        =   10
            Top             =   2670
            Width           =   8295
            _Version        =   393216
            _ExtentX        =   14631
            _ExtentY        =   11880
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
            SpreadDesigner  =   "LSN001.frx":0016
         End
         Begin EditLib.fpLongInteger fpLsnCapa 
            Height          =   360
            Left            =   4380
            TabIndex        =   8
            Top             =   2040
            Width           =   1095
            _Version        =   196608
            _ExtentX        =   1931
            _ExtentY        =   635
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   1
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
            Text            =   "0"
            MaxValue        =   "99999"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label15 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ǽ�"
            Height          =   210
            Left            =   300
            TabIndex        =   39
            Top             =   2130
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   3300
            TabIndex        =   38
            Top             =   1710
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   5640
            TabIndex        =   34
            Top             =   1740
            Width           =   975
         End
         Begin VB.Label lblLsnColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"LSN001.frx":19C9
            Height          =   675
            Left            =   6750
            TabIndex        =   9
            Top             =   1710
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� �ڵ��Ī"
            Height          =   210
            Left            =   3300
            TabIndex        =   31
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� ����                     ��"
            Height          =   210
            Left            =   2340
            TabIndex        =   30
            Top             =   2130
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   300
            TabIndex        =   29
            Top             =   1740
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���̸�"
            Height          =   210
            Left            =   300
            TabIndex        =   28
            Top             =   1290
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "LSN001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : LSN001
'   �� ��  �� �� : �� ���� ���
'
'   ��   ��   �� : 2007/08/24
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit




Private Sub Form_Click()
    If fraUpdate.Visible = True Then fraUpdate.Visible = False
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        fraUpdate.Visible = False
    End If
End Sub

Private Sub Label2_Click()
    fraUpdate.Visible = False
End Sub

Private Sub Frame1_Click()
    If fraUpdate.Visible = True Then fraUpdate.Visible = False
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0, 8805, 10060
    
    Me.Tag = "LOAD"
        With sprDisp
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With cboUpKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        'fraUpdate.Move 10000, 3090, 4845, 4095
        fraUpdate.ZOrder 0
        fraUpdate.Visible = False
            
        Call init_Form
        
    Me.Tag = ""
    
End Sub


Private Sub init_Form()
    
    txtUpLsnCD.Text = ""
    
    txtLsnNM.Text = ""
    txtUpLsnNM.Text = ""
    
    txtLsnCDNM.Text = ""
    txtUpLsnCDNM.Text = ""
    
    fpLsnCapa.value = 0
    fpUpLsnCapa.value = 0
    
'<< �߰����� : 2008.01.31
    txtDamim.Text = ""
    txtUpDamim.Text = ""
    
    txtBase_Class.Text = ""
    txtUpBase_Class.Text = ""
        
    sprDisp.MaxRows = 0
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    cboKaeyol.ListIndex = 0
    cboUpKaeyol.ListIndex = 0
    
    lblLsnColor.BackColor = &HFFFFFF
    lblUpLsnColor.BackColor = &HFFFFFF
    
End Sub




'>> ������ ��ȸ�ϱ�
Private Sub cmdFind_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    
    Dim sStr        As String
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Integer
    Dim nRec        As Long
    Dim nColor      As Long
    
    On Error GoTo ErrStmt
    
    sprDisp.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM,  "
    sStr = sStr & "         KAEYOL , DECODE(KAEYOL,'01','�ι�','02','�ڿ�') AS KAEYOL_NM, "
    sStr = sStr & "         LSNCAPA, "
    sStr = sStr & "         LSN_CL, "
    sStr = sStr & "         DAMIM, BASE_CLASS "
    sStr = sStr & "    FROM SDLSN01TB "
    sStr = sStr & "   WHERE ACID   = ? "
    sStr = sStr & "     AND KAEYOL = ? "
    sStr = sStr & "   ORDER BY KAEYOL, LSNCDNM "
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
    '>> �п�
    sTmp = Trim(basModule.SchCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    
    '>>
    sTmp = Trim(Right(cboKaeyol.Text, 30))
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprDisp.MaxRows = sprDisp.MaxRows + 1
                sprDisp.Row = sprDisp.MaxRows:      sprDisp.RowHeight(sprDisp.Row) = 15
                
                sprDisp.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("KAEYOL")) = False Then sTmp = Trim(.Fields("KAEYOL"))
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("KAEYOL_NM")) = False Then sTmp = Trim(.Fields("KAEYOL_NM"))
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                
                '<< �߰� : 2008.01.31
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("DAMIM")) = False Then sTmp = Trim(.Fields("DAMIM"))
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("BASE_CLASS")) = False Then sTmp = Trim(.Fields("BASE_CLASS"))
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                
                sprDisp.Col = sprDisp.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNCAPA")) = False Then nTmp = CDbl(Trim(.Fields("LSNCAPA")))
                    Call basFunction.Set_SprType_Numeric(sprDisp, 0, 0, 99999, "", nTmp)
                
                sprDisp.Col = sprDisp.Col + 1
                    nColor = &HFFFFFF
                        If IsNumeric(.Fields("LSN_CL")) = True Then nColor = CLng(.Fields("LSN_CL"))
                        sprDisp.Row2 = sprDisp.Row
                        sprDisp.Col2 = sprDisp.Col
                        sprDisp.BlockMode = True
                            sprDisp.BackColor = nColor
                            sprDisp.BackColorStyle = BackColorStyleUnderGrid
                        sprDisp.BlockMode = False
                
                sprDisp.Col = sprDisp.Col + 1
                    sprDisp.CellType = CellTypeButton
                    sprDisp.TypeButtonText = "����"
                
                .MoveNext
            Next nRec
            
            sprDisp.Row = 1:       sprDisp.Row2 = sprDisp.MaxRows
            sprDisp.Col = 1:       sprDisp.Col2 = sprDisp.MaxCols
            sprDisp.BlockMode = True
                sprDisp.Lock = True
                sprDisp.Protect = True
            sprDisp.BlockMode = False
            
            sprDisp.Row = 1:       sprDisp.Row2 = sprDisp.MaxRows
            sprDisp.Col = 1:       sprDisp.Col2 = 6
            sprDisp.BlockMode = True
                sprDisp.BackColor = &HFFFFFF
                sprDisp.BackColorStyle = BackColorStyleUnderGrid
            sprDisp.BlockMode = False
                
        End If
    End With
    
    MsgBox "���� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�� ��ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ��ȸ"
End Sub



Private Sub lblLsnColor_Click()

    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
        lblLsnColor.BackColor = .color
    End With
    
    Exit Sub
ErrStmt:

End Sub

Private Sub lblUpLsnColor_Click()

    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
        lblUpLsnColor.BackColor = .color
    End With
    
    Exit Sub
ErrStmt:

End Sub

Private Sub sprDisp_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sTmp    As String
    
    If Row < 1 Then Exit Sub
    
    With sprDisp
        If .MaxRows < 1 Then Exit Sub
        If .Tag = "" Then .Tag = "1"
    
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 8
        .BlockMode = True
            .BackColor = &HFFFFFF
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = 8
        .BlockMode = True
        .BackColor = basModule.SelectColor2
        .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        .Col = 1:           sTmp = Trim(.Text):     txtUpLsnCD.Text = sTmp
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtUpLsnNM.Text = sTmp
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtUpLsnCDNM.Text = sTmp
        
        .Col = .Col + 1:    sTmp = Trim(.Text)      ' �迭�ڵ�
            Select Case sTmp
                Case "01"
                    cboUpKaeyol.ListIndex = 0
                Case "02"
                    cboUpKaeyol.ListIndex = 1
            End Select
        
        .Col = .Col + 1     '<< skip                  �迭��
        
        '<< �߰� : 2008.01.31
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtUpDamim.Text = sTmp
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtUpBase_Class.Text = sTmp
        
        .Col = .Col + 1:    sTmp = Trim(.Text)
            fpUpLsnCapa.value = CLng(sTmp)
        
        .Col = .Col + 1
            lblUpLsnColor.BackColor = .BackColor
        
        .Tag = Trim(CStr(Row))
        
        fraUpdate.Visible = False
        If Col = .MaxCols Then         '<< ������ư Ŭ���ÿ�
            fraUpdate.Move 510, 4280, 4845, 5115
            fraUpdate.ZOrder 0
            fraUpdate.Visible = True
            
        End If
        
    End With
    
End Sub



'<< ������ ����ϱ�
Private Sub cmdSave_Click()
    
    Dim bRet        As Boolean
    
    Dim ni          As Long
    Dim nRec        As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    If Trim(txtLsnNM.Text) = "" Then
        MsgBox "���̸��� �����ϴ�.", vbExclamation + vbOKOnly, "�� ����ϱ�"
        Exit Sub
    End If
        
    On Error GoTo ErrStmt
    
    cmdSave.Enabled = False
        bRet = Save_Lsn_Data
        
    cmdSave.Enabled = True
    
    If bRet = True Then
        MsgBox "�� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�� ����ϱ�"
    Else
        MsgBox "�� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ����ϱ�"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ����ϱ�"
    On Error GoTo 0
    
End Sub

'<< insert �� ����.
Private Function Save_Lsn_Data() As Boolean
    Dim bRet        As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nTmp        As Double
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim sLsnCD      As String
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                

        
    sStr = ""
    sStr = sStr & "  SELECT GET_LSNCD AS LSNCD FROM DUAL "
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            If IsNull(.Fields("LSNCD")) = False Then
                sLsnCD = Trim(.Fields("LSNCD"))
            Else
                sLsnCD = "00001"
            End If
        End If
    End With
        
    Set DBRec = Nothing
    
        
    '<< INSERT
    sStr = ""
    sStr = sStr & "  INSERT INTO SDLSN01TB (ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS, LSNCAPA, LSN_CL) "
    sStr = sStr & "  VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ? )"

    '>> ACID
    sTmp = Trim(basModule.SchCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '>> LSNCD
    sTmp = Trim(sLsnCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '>> LSNNM
    sTmp = Trim(txtLsnNM.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    
    '>> LSNCDNM
    sTmp = Trim(txtLsnCDNM.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
        
    '>> �迭
    sTmp = Trim(Right(cboKaeyol.Text, 30))
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '<< �߰� : 2007.01.31
        ' ����
    sTmp = Trim(txtDamim.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        ' �� ����
    sTmp = Trim(txtBase_Class)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    
    
    '>> �� ����
    nTmp = fpLsnCapa.value
        Set DBParam = DBCmd.CreateParameter("LSNCAPA", adDouble, adParamInput, , nTmp): DBCmd.Parameters.Append DBParam
        
    '>> �÷�
    nTmp = lblLsnColor.BackColor
        Set DBParam = DBCmd.CreateParameter("LSN_CL", adDouble, adParamInput, , nTmp): DBCmd.Parameters.Append DBParam
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    
            
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
    
        With sprDisp
            .MaxRows = .MaxRows + 1
            .InsertRows 1, 1
            .Row = 1:       .RowHeight(.Row) = 15
            
            .Col = 1
                sTmp = sLsnCD
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
            .Col = .Col + 1
                sTmp = Trim(txtLsnNM.Text)
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
            .Col = .Col + 1
                sTmp = Trim(txtLsnCDNM.Text)
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            .Col = .Col + 1
                sTmp = " ": sTmp = Trim(Right(cboKaeyol.Text, 30))
                Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                
                Select Case sTmp
                    Case "01"
                        .Col = .Col + 1
                        sTmp = "�ι�":     Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Case "02"
                        .Col = .Col + 1
                        sTmp = "�ڿ�":     Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                End Select
            
            '>> �߰� : 2008.01.31
            .Col = .Col + 1
                sTmp = Trim(txtDamim.Text)
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
            .Col = .Col + 1
                sTmp = Trim(txtBase_Class)
                    Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
            
            
            .Col = .Col + 1:    nTmp = fpLsnCapa.value:                     Call basFunction.Set_SprType_Numeric(sprDisp, 0, 0, 99999, "", nTmp)
            
            .Col = .Col + 1
                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = lblLsnColor.BackColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            
            .Col = .Col + 1
                .CellType = CellTypeButton
                .TypeButtonText = "����"
                
        End With
        
        Save_Lsn_Data = True
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_Lsn_Data = False
    
End Function



'<< �� ���� ����
Private Sub cmdUpdate_Click()
    Dim bRet        As Boolean
    
    Dim ni          As Long
    Dim nRec        As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    If Trim(txtUpLsnCD.Text) = "" Then
        MsgBox "�� ������ �����ϴ�." & vbCrLf & _
               "��ȸ �� �ٽ� �����ϱ� ��ư�� Ŭ���ϼ���.", vbExclamation + vbOKOnly, "�� �����ϱ�"
        Exit Sub
    End If
    If Trim(txtUpLsnNM.Text) = "" Then
        MsgBox "���̸��� �����ϴ�.", vbExclamation + vbOKOnly, "�� �����ϱ�"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdUpdate.Enabled = False
        bRet = Update_Lsn_Data
        
    cmdUpdate.Enabled = True
    
    If bRet = True Then
        MsgBox "�� ���� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�� �����ϱ�"
    Else
        MsgBox "�� ���� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� �����ϱ�"
    End If
    
    fraUpdate.Visible = False
    Exit Sub
ErrStmt:
    fraUpdate.Visible = False
    MsgBox "�� ���� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� �����ϱ�"
    On Error GoTo 0
    
End Sub

'<< update �� ����.
Private Function Update_Lsn_Data() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long
    Dim nRow        As Long
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
               
        
    sStr = ""
    sStr = sStr & " UPDATE SDLSN01TB "
    sStr = sStr & "    SET LSNNM      = ? ,"
    sStr = sStr & "        LSNCDNM    = ? ,"
    sStr = sStr & "        KAEYOL     = ? ,"
    
    sStr = sStr & "        DAMIM      = ? ,"
    sStr = sStr & "        BASE_CLASS = ?, "
    
    sStr = sStr & "        LSNCAPA    = ? ,"
    sStr = sStr & "        LSN_CL     = ?  "
    sStr = sStr & "  WHERE ACID    = ? "
    sStr = sStr & "    AND LSNCD   = ? "

    '>> LSNNM
    sTmp = Trim(txtUpLsnNM.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '>> LSNCDNM
    sTmp = Trim(txtUpLsnCDNM.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCDNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '>> �迭
    sTmp = Trim(Right(cboUpKaeyol.Text, 30))
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '<< �߰� : 2008.01.31
        '>> ����
    sTmp = Trim(txtUpDamim.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("DAMIM", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        '>> ���ǽ�
    sTmp = Trim(txtUpBase_Class)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("BASE_CLASS", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
        
    '>> �� ����
    nTmp = fpUpLsnCapa
        Set DBParam = DBCmd.CreateParameter("LSNCAPA", adDouble, adParamInput, , nTmp): DBCmd.Parameters.Append DBParam
        
    '>> �÷�
    nTmp = lblUpLsnColor.BackColor
        Set DBParam = DBCmd.CreateParameter("LSN_CL", adDouble, adParamInput, , nTmp): DBCmd.Parameters.Append DBParam
        
    '>> ACID
    sTmp = Trim(basModule.SchCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '>> �� �ڵ�
    sTmp = Trim(txtUpLsnCD.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
        sTmp = Trim(txtUpLsnCD.Text)
        nRow = 0
        With sprDisp
            For nRec = 1 To .MaxRows Step 1
                .Row = nRec
                .Col = 1
                    
                If StrComp(Trim(.Text), sTmp, vbTextCompare) = 0 Then
                    nRow = .Row
                    Exit For
                End If
            Next nRec
            
            
            If nRow > 0 Then
                
                .Row = nRow
                
                .Col = 1
                    sTmp = " ": sTmp = Trim(txtUpLsnCD.Text)
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                .Col = .Col + 1
                    sTmp = " ": sTmp = Trim(txtUpLsnNM.Text)
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                .Col = .Col + 1
                    sTmp = " ": sTmp = Trim(txtUpLsnCDNM.Text)
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                .Col = .Col + 1
                    sTmp = " ": sTmp = Trim(Right(cboUpKaeyol.Text, 30))
                        Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Select Case sTmp
                        Case "01"
                            .Col = .Col + 1
                            sTmp = "�ι�":     Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                        Case "02"
                            .Col = .Col + 1
                            sTmp = "�ڿ�":     Call basFunction.Set_SprType_Text(sprDisp, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    End Select
                
                '>> �߰� : 2008.01.31
                .Col = .Col + 1
                    sTmp = " ": sTmp = Trim(txtUpDamim.Text)
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                .Col = .Col + 1
                    sTmp = " ": sTmp = Trim(txtUpBase_Class.Text)
                        Call basFunction.Set_SprType_Text(sprDisp, "center", "left", basFunction.LenKor(sTmp), sTmp)
                
                
                .Col = .Col + 1
                    nTmp = 0:   nTmp = fpUpLsnCapa.value
                        Call basFunction.Set_SprType_Numeric(sprDisp, 0, 0, 99999, "", nTmp)
                
                .Col = .Col + 1
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = lblUpLsnColor.BackColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                
                .Col = .Col + 1
                    .CellType = CellTypeButton
                    .TypeButtonText = "����"
            
                Update_Lsn_Data = True
                
            End If
        End With
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Update_Lsn_Data = False
    
End Function




'<< ����
Private Sub cmdDelete_Click()
    Dim bRet        As Boolean
    
    Dim ni          As Long
    Dim nRec        As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If Trim(txtUpLsnCD.Text) = "" Then
        MsgBox "�� ������ �����ϴ�." & vbCrLf & _
               "��ȸ �� �ٽ� �����ϱ� ��ư�� Ŭ���ϼ���.", vbExclamation + vbOKOnly, "�� �����ϱ�"
        Exit Sub
    End If
    
    If StrComp(InputBox("�����ڵ带 ��������.", "�� �����ϱ�", ""), "DEL", vbTextCompare) <> 0 Then
        MsgBox "�����ڿ��� �����ϼ���.", vbExclamation + vbOKOnly, "�� �����ϱ�"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdDelete.Enabled = False
        bRet = Delete_Lsn_Data
        
    cmdDelete.Enabled = True
    
    If bRet = True Then
        MsgBox "�� ���� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�� �����ϱ�"
    Else
        MsgBox "�� ���� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� �����ϱ�"
    End If
    
    fraUpdate.Visible = False
    Exit Sub
ErrStmt:
    fraUpdate.Visible = False
    MsgBox "�� ���� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� �����ϱ�"
    On Error GoTo 0
    
End Sub

'<< update �� ����.
Private Function Delete_Lsn_Data() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim nRow        As Long
    Dim nRec        As Long
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
    sStr = ""
    sStr = sStr & " DELETE SDLSN01TB "
    sStr = sStr & "  WHERE ACID  = ? "
    sStr = sStr & "    AND LSNCD = ? "

    '>> ACID
    sTmp = Trim(basModule.SchCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    '>> �� �ڵ�
    sTmp = Trim(txtUpLsnCD.Text)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    
    If nExe = 1 Then
        sTmp = Trim(txtUpLsnCD.Text)
        nRow = 0
        With sprDisp
            For nRec = .MaxRows To 1 Step -1
                .Row = nRec
                .Col = 1
                
                If StrComp(Trim(.Text), sTmp, vbTextCompare) = 0 Then
                    nRow = .Row
                    
                    .DeleteRows nRow, 1
                    .MaxRows = .MaxRows - 1
                    
                    
                    Delete_Lsn_Data = True
                    Exit For
                End If
            Next nRec
            
        End With
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_Lsn_Data = False
    
End Function


