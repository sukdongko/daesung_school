VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form TMR025 
   Caption         =   "�ð�ǥ ����� >> �̵����� �ð�ǥ ��� CP"
   ClientHeight    =   11670
   ClientLeft      =   240
   ClientTop       =   1995
   ClientWidth     =   15990
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11670
   ScaleWidth      =   15990
   WindowState     =   2  '�ִ�ȭ
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   360
      TabIndex        =   13
      Top             =   7770
      Width           =   1980
   End
   Begin VB.CommandButton cmdSelGwamok 
      Caption         =   "���ó��� �ݿ��ϱ�"
      Height          =   330
      Left            =   5310
      TabIndex        =   12
      Top             =   3060
      Width           =   1980
   End
   Begin FPSpread.vaSpread sprClass 
      Height          =   2775
      Left            =   5220
      TabIndex        =   11
      Top             =   120
      Width           =   12405
      _Version        =   393216
      _ExtentX        =   21881
      _ExtentY        =   4895
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
      MaxRows         =   4
      SpreadDesigner  =   "TMR025.frx":0000
   End
   Begin VB.CommandButton cmdControlMvBan 
      Caption         =   "�̵��� ����"
      Height          =   330
      Left            =   2910
      TabIndex        =   10
      Top             =   3060
      Width           =   1110
   End
   Begin VB.CommandButton cmdControlBan 
      Caption         =   "�� ����"
      Height          =   330
      Left            =   1590
      TabIndex        =   9
      Top             =   3060
      Width           =   1110
   End
   Begin VB.CommandButton cmdFindBan 
      Caption         =   "�� ��ȸ"
      Height          =   330
      Left            =   300
      TabIndex        =   8
      Top             =   3060
      Width           =   1110
   End
   Begin FPSpread.vaSpread sprLsn 
      Height          =   4185
      Left            =   5250
      TabIndex        =   1
      Top             =   3450
      Width           =   12435
      _Version        =   393216
      _ExtentX        =   21934
      _ExtentY        =   7382
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
      MaxCols         =   19
      SpreadDesigner  =   "TMR025.frx":2FF3
   End
   Begin FPSpread.vaSpread sprBan 
      Height          =   4185
      Left            =   300
      TabIndex        =   7
      Top             =   3450
      Width           =   4905
      _Version        =   393216
      _ExtentX        =   8652
      _ExtentY        =   7382
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
      SpreadDesigner  =   "TMR025.frx":4DA9
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��ȸ"
      Height          =   465
      Left            =   270
      TabIndex        =   6
      Top             =   120
      Width           =   1725
   End
   Begin VB.ComboBox cboKaeyol 
      Height          =   300
      Left            =   3030
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   4
      Top             =   570
      Width           =   1005
   End
   Begin VB.ComboBox cboExmType 
      Height          =   300
      Left            =   3030
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
   Begin FPSpread.vaSpread sprData 
      Height          =   3435
      Left            =   270
      TabIndex        =   0
      Top             =   8340
      Width           =   17415
      _Version        =   393216
      _ExtentX        =   30718
      _ExtentY        =   6059
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
      MaxCols         =   28
      SpreadDesigner  =   "TMR025.frx":66E1
   End
   Begin VB.Label Label2 
      Caption         =   "�� ������ ���� �����Ͻ� �� �������� �ش�ݿ� ������ ���ð���4���� �����ϼ���."
      Height          =   195
      Left            =   7440
      TabIndex        =   15
      Top             =   2940
      Width           =   10185
   End
   Begin VB.Label lblStatus 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7440
      TabIndex        =   14
      Top             =   3180
      Width           =   10185
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�迭"
      Height          =   210
      Left            =   1980
      TabIndex        =   5
      Top             =   615
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��/������"
      Height          =   210
      Left            =   1980
      TabIndex        =   3
      Top             =   225
      Width           =   975
   End
End
Attribute VB_Name = "TMR025"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM025
'   �� ��  �� �� : �̵����� �ð�ǥ ���
'
'   ��   ��   �� : 2008/01/04
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private Const nRowHeight = 15




Private Sub Form_Load()
    
    Me.Move 0, 0, 15195, 9630
    
    Me.Tag = "LOAD"
        With sprData
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxRows = 0
        End With
        
        With sprLsn
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxRows = 0
        End With
                
        With sprBan
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxRows = 0
        End With
                
        With sprClass
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxCols = 0
        End With
                
        With cboExmType
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "������" & Space(30) & "0"
            .AddItem "������" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
                
    Me.Tag = ""
    
    lblStatus.Caption = ""
    
End Sub

Private Sub cboKaeyol_Click()
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"
            With sprLsn
                .Row = SpreadHeader + 1
                .Col = 4:           .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "�ѱ�"
                .Col = .Col + 1:    .Text = "�����"
                
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "��ġ"
                .Col = .Col + 1:    .Text = "�繮"
                .Col = .Col + 1:    .Text = "����"
                
                .Col = .Col + 1:    .Text = "����"
                
                .MaxRows = 0
                
            End With
            
            With sprData
                .Row = SpreadHeader + 1
                .Col = 13:          .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "�ѱ�"
                .Col = .Col + 1:    .Text = "�����"
                
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "����"
                .Col = .Col + 1:    .Text = "��ġ"
                .Col = .Col + 1:    .Text = "�繮"
                .Col = .Col + 1:    .Text = "����"
                
                .Col = .Col + 1:    .Text = "����"
                
                .MaxRows = 0
                
            End With
            
        Case "02"
            With sprLsn
                .Row = SpreadHeader + 1
                .Col = 4:           .Text = "��1"
                .Col = .Col + 1:    .Text = "ȭ1"
                .Col = .Col + 1:    .Text = "��1"
                .Col = .Col + 1:    .Text = "��1"
                .Col = .Col + 1:    .Text = "��2"
                
                .Col = .Col + 1:    .Text = "ȭ2"
                .Col = .Col + 1:    .Text = "��2"
                .Col = .Col + 1:    .Text = "��2"
                .Col = .Col + 1:    .Text = ""
                .Col = .Col + 1:    .Text = ""
                
                .Col = .Col + 1:    .Text = ""
                
                .MaxRows = 0
                
            End With
            
            With sprData
                .Row = SpreadHeader + 1
                .Col = 13:          .Text = "��1"
                .Col = .Col + 1:    .Text = "ȭ1"
                .Col = .Col + 1:    .Text = "��1"
                .Col = .Col + 1:    .Text = "��1"
                .Col = .Col + 1:    .Text = "��2"
                
                .Col = .Col + 1:    .Text = "ȭ2"
                .Col = .Col + 1:    .Text = "��2"
                .Col = .Col + 1:    .Text = "��2"
                .Col = .Col + 1:    .Text = ""
                .Col = .Col + 1:    .Text = ""
                
                .Col = .Col + 1:    .Text = ""
                
                .MaxRows = 0
                
            End With
            
            
    End Select
    
End Sub




Private Sub cmdControlBan_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub

Private Sub cmdControlMvBan_Click()
    Load LSN002
    LSN002.Show
    LSN002.ZOrder 0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Unload LSN001
    Unload LSN002
    
End Sub


Private Sub cmdFindBan_Click()
    Call Find_Lsn_Data              '< �� ������ȸ
    MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ȸ"
    
End Sub


'>> ������ ��ȸ
Private Sub cmdFind_Click()
    
    Call Find_Lsn_Data              '< �� ������ȸ
        
    Call Find_Lsn_To_STD_TOT        '< �ݺ��� �հ賻��
    Call Find_Gwamok_to_STD_TOT     '< ���� �հ賻��
    
    Call Find_STD_Data              '< �л���ȸ
    
    MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ȸ"

End Sub

'## �� ������ȸ
Private Sub Find_Lsn_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nColor      As Long
    
    Dim sFieldNM    As String
    Dim sCboBoxList As String
    
    sprBan.MaxRows = 0
    sprClass.MaxCols = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM, LSNCAPA, SEL_OK, LSN_CL"
    sStr = sStr & "    FROM (SELECT *"
    sStr = sStr & "            From SDLSN01TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "           ORDER BY LSNCDNM"
    sStr = sStr & "          )"
    sStr = sStr & "  Union All"
    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM, LSNCAPA, SEL_OK, LSN_CL"
    sStr = sStr & "    FROM (SELECT *"
    sStr = sStr & "            From SDLSN02TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "           ORDER BY LSNCDNM"
    sStr = sStr & "          )"
    
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
    
'    '>> �п�
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
                sprBan.MaxRows = sprBan.MaxRows + 1
                sprBan.Row = sprBan.MaxRows:    sprBan.RowHeight(sprBan.Row) = nRowHeight


                sprBan.Col = 1
                    sTmp = " ":     If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprBan, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                    '~1. ���ڵ� ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        sprClass.MaxCols = sprClass.MaxCols + 1
                        sprClass.Col = sprClass.MaxCols
                        
                        sprClass.Row = SpreadHeader
                            sprClass.Text = sTmp
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    
                sprBan.Col = sprBan.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprBan, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                    '~1. �ݸ� ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        sprClass.Col = sprClass.MaxCols
                        
                        sprClass.Row = SpreadHeader + 1
                            sprClass.Text = sTmp
                            
                        sprClass.Row = 1
                            
                            sprClass.CellType = CellTypeComboBox
                            sprClass.TypeComboBoxClear sprClass.Col, sprClass.Row
                            sCboBoxList = "����" & Chr$(9) & _
                                          "A" & Chr$(9) & _
                                          "B" & Chr$(9) & _
                                          "C" & Chr$(9) & _
                                          "����" & Chr$(9)
                            sprClass.TypeComboBoxList = sCboBoxList
                            sprClass.TypeComboBoxCurSel = 0
                            
                            sprClass.TypeHAlign = TypeHAlignCenter
                            sprClass.TypeVAlign = TypeVAlignCenter
                    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    
                sprBan.Col = sprBan.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprBan, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprBan.SetCellBorder sprBan.Col, sprBan.Row, sprBan.Col, sprBan.Row, 2, basModule.SectionColor1, CellBorderStyleSolid

                sprBan.Col = sprBan.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("LSNCAPA")) = True Then nTmp = CLng(.Fields("LSNCAPA"))
                        Call basFunction.Set_SprType_Numeric(sprBan, 0, -99999, 99999, ",", nTmp)
                        
                sprBan.Col = sprBan.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("SEL_OK")) = True Then nTmp = CLng(.Fields("SEL_OK"))
                        Call basFunction.Set_SprType_Numeric(sprBan, 0, -99999, 99999, ",", nTmp)

                sprBan.Col = sprBan.Col + 1
                    nColor = &HFFFFFF
                        If IsNumeric(.Fields("LSN_CL")) = True Then nColor = CLng(.Fields("LSN_CL"))
                        sprBan.Row2 = sprBan.Row
                        sprBan.Col2 = sprBan.Col
                        sprBan.BlockMode = True
                            sprBan.BackColor = nColor
                            sprBan.BackColorStyle = BackColorStyleUnderGrid
                        sprBan.BlockMode = False
                        
                sprBan.Col = sprBan.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprBan)
                    sprBan.Value = 0
                
                
                .MoveNext       '<< �����׸�
                
            Next nRec
        End If
        
        With sprBan
'            .Row = 1:       .Row2 = .MaxRows
'            .Col = 1:       .Col2 = .MaxCols
'            .BlockMode = True
'                .BackColor = basModule.WhiteColor
'                .BackColorStyle = BackColorStyleUnderGrid
'            .BlockMode = False
            
            .ColsFrozen = 3
            
        '>> spread lock
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ��ȸ"


End Sub


'## �л��� ��û���� ��ȸ
Private Sub Find_STD_Data()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim sFieldNM    As String
    
    sprData.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, EXMID, STDNM, "
    sStr = sStr & "         EXMTYPE, EXMTYPE_NM,"
    sStr = sStr & "         GAEYUL_CD, GAEYUL,"
    sStr = sStr & "         SEL1, SEL2, SEL3, SEL4, SEL5 ,"
    sStr = sStr & "         SEL6, SEL7, SEL8, SEL9, SEL10,"
    sStr = sStr & "         SEL11,"
    
    sStr = sStr & "         SEL_X2,"
    sStr = sStr & "         SEL_N1, SEL_N2, SEL_N3, SEL_N4,"
    sStr = sStr & "         SEL_CLASS,"
    sStr = sStr & "         SEL_CLASS_NM,"
    sStr = sStr & "         CL_CLOSE,"
    sStr = sStr & "         GWA_BAN1, GWA_BAN2, GWA_BAN3, GWA_BAN4"
    sStr = sStr & "    FROM (SELECT SCHNO, EXMID, STDNM,"
    sStr = sStr & "                 EXMTYPE, DECODE(EXMTYPE,'0','��','��') AS EXMTYPE_NM,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                     '01'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                     '02'"
    sStr = sStr & "                 END END GAEYUL_CD,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                     '��Ž'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                     '��Ž'"
    sStr = sStr & "                 END END GAEYUL,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN"
    sStr = sStr & "                     '��1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL1,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN"
    sStr = sStr & "                     'ȭ1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL2,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN"
    sStr = sStr & "                     '��1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL3,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN"
    sStr = sStr & "                     '�ѱ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN"
    sStr = sStr & "                     '��1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL4,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN"
    sStr = sStr & "                     '�����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN"
    sStr = sStr & "                     '��2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL5,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN"
    sStr = sStr & "                     'ȭ2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL6,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN"
    sStr = sStr & "                     '��2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL7,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN"
    sStr = sStr & "                     '��ġ'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN"
    sStr = sStr & "                     '��2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL8,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN"
    sStr = sStr & "                     '�繮'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL9,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL10,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL11,"
    
    sStr = sStr & "              /* ��2�ܱ��� & ���� */"
    sStr = sStr & "                      CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '���Ͼ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '�����ĳ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߱���'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '������'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ�����'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '��������'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "              /* ��� */"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                     '���'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N1,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                     '����'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N2,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                     '��Ž'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N3,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                     '��Ž'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N4,"
    
    sStr = sStr & "                 SEL_CLASS, GET_LSNNM(ACID, SEL_CLASS) AS SEL_CLASS_NM,"
    sStr = sStr & "                 CL_CLOSE,"
    sStr = sStr & "                 GET_LSNNM(ACID, GWA_BAN1) AS GWA_BAN1,"
    sStr = sStr & "                 GET_LSNNM(ACID, GWA_BAN2) AS GWA_BAN2,"
    sStr = sStr & "                 GET_LSNNM(ACID, GWA_BAN3) AS GWA_BAN3,"
    sStr = sStr & "                 GET_LSNNM(ACID, GWA_BAN4) AS GWA_BAN4"
    sStr = sStr & "            FROM CLTTL01TB"
    sStr = sStr & "           WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    
    
    
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "0"
            sStr = sStr & "         AND EXMTYPE = '0' "
        Case "1"
            sStr = sStr & "         AND EXMTYPE = '1' "
        Case Else
            ' NO ACTION
    End Select
    
    sStr = sStr & "        )"
    sStr = sStr & "    WHERE EXMID > ' ' "
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"
            sStr = sStr & " AND GAEYUL_CD = '01' "
        Case "02"
            sStr = sStr & " AND GAEYUL_CD = '02' "
        Case Else
            ' NO ACTION
    End Select
    sStr = sStr & "   ORDER BY SEL_CLASS, GAEYUL_CD, EXMID, STDNM"
    
    
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
    
'    '>> �п�
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
                sprData.MaxRows = sprData.MaxRows + 1
                sprData.Row = sprData.MaxRows ':      sprData.RowHeight(sprData.Row) = nRowHeight

                sprData.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor1, CellBorderStyleSolid


                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMTYPE")) = False Then sTmp = Trim(.Fields("EXMTYPE"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMTYPE_NM")) = False Then sTmp = Trim(.Fields("EXMTYPE_NM"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)

                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS")) = False Then sTmp = Trim(.Fields("SEL_CLASS"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS_NM")) = False Then sTmp = Trim(.Fields("SEL_CLASS_NM"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                For ni = 1 To 4 Step 1
                    sFieldNM = ""

                    sFieldNM = "GWA_BAN" & Trim(CStr(ni))
                    sprData.Col = sprData.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni
                
                sprData.Col = sprData.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprData)
                    sprData.Value = 0

                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                For ni = 1 To 11 Step 1
                    sFieldNM = ""

                    sFieldNM = "SEL" & Trim(CStr(ni))
                    sprData.Col = sprData.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni

                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("SEL_X2")) = False Then sTmp = Trim(.Fields("SEL_X2"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)

                
                .MoveNext       '<< �����׸�
                
            Next nRec
        End If
        
        With sprData
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .ColsFrozen = 3
            
        '>> spread lock
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�л� ��û���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"
    
End Sub

'## �ݺ� �����û���� �հ��ο�
Private Function Find_Lsn_To_STD_TOT() As Long

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Integer
    Dim siTem       As String
    
    Dim nRet        As Long
    
    On Error GoTo ErrStmt
    
    nRet = 0
    sprLsn.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, INWON_STAT, "
    sStr = sStr & "         S_LSN,"
    sStr = sStr & "         SEL1 ,"
    sStr = sStr & "         SEL2 ,"
    sStr = sStr & "         SEL3 ,"
    sStr = sStr & "         SEL4 ,"
    sStr = sStr & "         SEL5 ,"
    sStr = sStr & "         SEL6 ,"
    sStr = sStr & "         SEL7 ,"
    sStr = sStr & "         SEL8 ,"
    sStr = sStr & "         SEL9 ,"
    sStr = sStr & "         SEL10,"
    sStr = sStr & "         SEL11,"
    
    sStr = sStr & "         SEL_X2,"
    
    sStr = sStr & "         SEL_N1,"
    sStr = sStr & "         SEL_N2,"
    sStr = sStr & "         SEL_N3,"
    sStr = sStr & "         SEL_N4,"
    
    sStr = sStr & "         KAEYOL, "
    sStr = sStr & "         DECODE(KAEYOL,'01','�ι�',"
    sStr = sStr & "                       '02','�ڿ� J') AS KAEYOL_NM"
    
    sStr = sStr & "    FROM (SELECT LSNCD,"
    sStr = sStr & "                 GET_LSNNM(ACID, LSNCD) AS LSNNM,"
    
    sStr = sStr & "                 COUNT(CL_CLOSE) AS INWON_STAT,                      /* �۾��Ϸ� �� �л� */"
    
    sStr = sStr & "                 COUNT(LSNCD) AS S_LSN,"
    sStr = sStr & "                 SUM(SEL1 ) AS SEL1 ,"
    sStr = sStr & "                 SUM(SEL2 ) AS SEL2 ,"
    sStr = sStr & "                 SUM(SEL3 ) AS SEL3 ,"
    sStr = sStr & "                 SUM(SEL4 ) AS SEL4 ,"
    sStr = sStr & "                 SUM(SEL5 ) AS SEL5 ,"
    sStr = sStr & "                 SUM(SEL6 ) AS SEL6 ,"
    sStr = sStr & "                 SUM(SEL7 ) AS SEL7 ,"
    sStr = sStr & "                 SUM(SEL8 ) AS SEL8 ,"
    sStr = sStr & "                 SUM(SEL9 ) AS SEL9 ,"
    sStr = sStr & "                 SUM(SEL10) AS SEL10,"
    sStr = sStr & "                 SUM(SEL11) AS SEL11,"
    
    sStr = sStr & "                 SUM(SEL_X2) AS SEL_X2,"

    sStr = sStr & "                 SUM(SEL_N1) AS SEL_N1,"
    sStr = sStr & "                 SUM(SEL_N2) AS SEL_N2,"
    sStr = sStr & "                 SUM(SEL_N3) AS SEL_N3,"
    sStr = sStr & "                 SUM(SEL_N4) AS SEL_N4,"
    
    sStr = sStr & "                 MAX(GAEYUL_CD) AS KAEYOL"
    
    sStr = sStr & "           FROM (SELECT ACID, LSNCD, "
    sStr = sStr & "                        GAEYUL_CD,"
    
    sStr = sStr & "                        SEL1 ,"
    sStr = sStr & "                        SEL2 ,"
    sStr = sStr & "                        SEL3 ,"
    sStr = sStr & "                        SEL4 ,"
    sStr = sStr & "                        SEL5 ,"
    sStr = sStr & "                        SEL6 ,"
    sStr = sStr & "                        SEL7 ,"
    sStr = sStr & "                        SEL8 ,"
    sStr = sStr & "                        SEL9 ,"
    sStr = sStr & "                        SEL10,"
    sStr = sStr & "                        SEL11,"
    
    sStr = sStr & "                        SEL_X2,"
    
    sStr = sStr & "                        SEL_N1,"
    sStr = sStr & "                        SEL_N2,"
    sStr = sStr & "                        SEL_N3,"
    sStr = sStr & "                        SEL_N4,"
    
    sStr = sStr & "                        CL_CLOSE "
    
    sStr = sStr & "                  FROM (SELECT "
    sStr = sStr & "                               SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* ��Ž, ��Ž �и� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    
    sStr = sStr & "                           /* ��2�ܱ��� & ���� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '���Ͼ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '�����ĳ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߱���'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '������'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻����'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ�����'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '��������'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                        ''"
    sStr = sStr & "                               END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "                           /* ��� */"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                                   '���'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N1,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                                       '����'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N2,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                                       '��Ž'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N3,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                                       '��Ž'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N4,"
    
    sStr = sStr & "                               CL_CLOSE "
    
    sStr = sStr & "                          FROM CLTTL01TB"
    sStr = sStr & "                         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        )"
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"
            sStr = sStr & "            WHERE GAEYUL_CD = '01' "
        Case "02"
            sStr = sStr & "            WHERE GAEYUL_CD = '02' "
        Case Else
            ' NO ACTION
    End Select
    
    sStr = sStr & "                   )"
    sStr = sStr & "              GROUP BY LSNCD"
    sStr = sStr & "              HAVING LSNCD"
    sStr = sStr & "                  IN (SELECT LSNCD"
    sStr = sStr & "                        FROM SDLSN01TB"
    sStr = sStr & "                       WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"
            sStr = sStr & "                 AND KAEYOL = '01' "
        Case "02"
            sStr = sStr & "                 AND KAEYOL = '02' "
        Case Else
            ' NO ACTION
    End Select
    sStr = sStr & "                     )"
    sStr = sStr & "           )"
    sStr = sStr & "      ORDER BY LSNNM "
    
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
            
                nRet = nRet + 1
                
                sprLsn.MaxRows = sprLsn.MaxRows + 1
                sprLsn.Row = sprLsn.MaxRows:        sprLsn.RowHeight(sprLsn.Row) = nRowHeight
                
                sprLsn.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprLsn.Col = sprLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> ���ο�
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
                
                '<< �ι��ڿ� ���� : 8 ���� >>
                For nCol = 1 To 8 Step 1
                    sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '��Ž�� 9~11
                        For nCol = 9 To 11 Step 1
                            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                            
                        Next nCol
                        
                    Case "02"
                        '��Ž�� COLUMN�� �̵�
                        For nCol = 9 To 11 Step 1
                            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                            If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                        Next nCol
                End Select
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> ��2����
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_X2")) = False Then
                        nTmp = CDbl(.Fields("SEL_X2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N1")) = False Then
                        nTmp = CDbl(.Fields("SEL_N1"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N2")) = False Then
                        nTmp = CDbl(.Fields("SEL_N2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N3")) = False Then
                        nTmp = CDbl(.Fields("SEL_N3"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> Ž
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N4")) = False Then
                        nTmp = CDbl(.Fields("SEL_N4"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                
                .MoveNext       '<< �����׸�
                
            Next nRec
            
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.BackColor = basModule.WhiteColor
                sprLsn.BackColorStyle = BackColorStyleUnderGrid
            sprLsn.BlockMode = False

            sprLsn.ColsFrozen = 5
            
        '>> spread lock
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.Lock = True
                sprLsn.Protect = True
            sprLsn.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Find_Lsn_To_STD_TOT = nRet
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�ݺ� ������û���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ݺ� ������û���� ��ȸ"
    
    Find_Lsn_To_STD_TOT = nRet
End Function





'## ��ü ���� �л���
Private Function Find_Gwamok_to_STD_TOT() As Long

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Integer
    Dim siTem       As String
    
    Dim nRet        As Long
    
    On Error GoTo ErrStmt
    
    nRet = 0
    
    sStr = ""
    sStr = sStr & "  SELECT INWON_STAT, "
    sStr = sStr & "         S_LSN,"
    sStr = sStr & "         SEL1 ,"
    sStr = sStr & "         SEL2 ,"
    sStr = sStr & "         SEL3 ,"
    sStr = sStr & "         SEL4 ,"
    sStr = sStr & "         SEL5 ,"
    sStr = sStr & "         SEL6 ,"
    sStr = sStr & "         SEL7 ,"
    sStr = sStr & "         SEL8 ,"
    sStr = sStr & "         SEL9 ,"
    sStr = sStr & "         SEL10,"
    sStr = sStr & "         SEL11,"
    
    sStr = sStr & "         SEL_X2,"
    
    sStr = sStr & "         SEL_N1,"
    sStr = sStr & "         SEL_N2,"
    sStr = sStr & "         SEL_N3,"
    sStr = sStr & "         SEL_N4"
    
    sStr = sStr & "    FROM (SELECT COUNT(CL_CLOSE) AS INWON_STAT,                      /* �۾��Ϸ� �� �л� */"
    
    sStr = sStr & "                 COUNT(LSNCD) AS S_LSN,"
    sStr = sStr & "                 SUM(SEL1 ) AS SEL1 ,"
    sStr = sStr & "                 SUM(SEL2 ) AS SEL2 ,"
    sStr = sStr & "                 SUM(SEL3 ) AS SEL3 ,"
    sStr = sStr & "                 SUM(SEL4 ) AS SEL4 ,"
    sStr = sStr & "                 SUM(SEL5 ) AS SEL5 ,"
    sStr = sStr & "                 SUM(SEL6 ) AS SEL6 ,"
    sStr = sStr & "                 SUM(SEL7 ) AS SEL7 ,"
    sStr = sStr & "                 SUM(SEL8 ) AS SEL8 ,"
    sStr = sStr & "                 SUM(SEL9 ) AS SEL9 ,"
    sStr = sStr & "                 SUM(SEL10) AS SEL10,"
    sStr = sStr & "                 SUM(SEL11) AS SEL11,"
    
    sStr = sStr & "                 SUM(SEL_X2) AS SEL_X2,"

    sStr = sStr & "                 SUM(SEL_N1) AS SEL_N1,"
    sStr = sStr & "                 SUM(SEL_N2) AS SEL_N2,"
    sStr = sStr & "                 SUM(SEL_N3) AS SEL_N3,"
    sStr = sStr & "                 SUM(SEL_N4) AS SEL_N4"
    
    sStr = sStr & "           FROM (SELECT LSNCD, "
    sStr = sStr & "                        GAEYUL_CD,"
    
    sStr = sStr & "                        SEL1 ,"
    sStr = sStr & "                        SEL2 ,"
    sStr = sStr & "                        SEL3 ,"
    sStr = sStr & "                        SEL4 ,"
    sStr = sStr & "                        SEL5 ,"
    sStr = sStr & "                        SEL6 ,"
    sStr = sStr & "                        SEL7 ,"
    sStr = sStr & "                        SEL8 ,"
    sStr = sStr & "                        SEL9 ,"
    sStr = sStr & "                        SEL10,"
    sStr = sStr & "                        SEL11,"
    
    sStr = sStr & "                        SEL_X2,"
    
    sStr = sStr & "                        SEL_N1,"
    sStr = sStr & "                        SEL_N2,"
    sStr = sStr & "                        SEL_N3,"
    sStr = sStr & "                        SEL_N4,"
    
    sStr = sStr & "                        CL_CLOSE "
    
    sStr = sStr & "                  FROM (SELECT "
    sStr = sStr & "                               SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* ��Ž, ��Ž �и� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    
    sStr = sStr & "                           /* ��2�ܱ��� & ���� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '���Ͼ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '�����ĳ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߱���'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '������'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻����'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ�����'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '��������'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                        ''"
    sStr = sStr & "                               END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "                           /* ��� */"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                                   '���'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N1,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                                       '����'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N2,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                                       '��Ž'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N3,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                                       '��Ž'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N4,"
    
    sStr = sStr & "                               CL_CLOSE "
    
    sStr = sStr & "                          FROM CLTTL01TB"
    sStr = sStr & "                         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        )"
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"
            sStr = sStr & "            WHERE GAEYUL_CD = '01' "
        Case "02"
            sStr = sStr & "            WHERE GAEYUL_CD = '02' "
        Case Else
            ' NO ACTION
    End Select
    
    sStr = sStr & "                   )"
    sStr = sStr & "                )"
    
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            sprLsn.MaxRows = sprLsn.MaxRows + 1
            sprLsn.InsertRows 1, 1
            sprLsn.Row = 1:     sprLsn.RowHeight(sprLsn.Row) = nRowHeight
            
            sprLsn.SetCellBorder 1, sprLsn.Row, sprLsn.MaxCols, sprLsn.Row, 8, basModule.SectionColor1, CellBorderStyleSolid
                
                
            sprLsn.Col = 1
                sTmp = " "
                    Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
            
            sprLsn.Col = sprLsn.Col + 1
                sTmp = "��    ��"
                    Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprLsn.ForeColor = basModule.SectionColor1
                    sprLsn.TypeHAlign = TypeHAlignCenter
                
            'sprLsn.ForeColor = &H0
            sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
            
            '>> ���ο�
            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
            sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '<< �ι��ڿ� ���� : 8 ���� >>
                For nCol = 1 To 11 Step 1
                    sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> ��2����
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_X2")) = False Then
                        nTmp = CDbl(.Fields("SEL_X2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N1")) = False Then
                        nTmp = CDbl(.Fields("SEL_N1"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N2")) = False Then
                        nTmp = CDbl(.Fields("SEL_N2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> ��
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N3")) = False Then
                        nTmp = CDbl(.Fields("SEL_N3"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> Ž
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N4")) = False Then
                        nTmp = CDbl(.Fields("SEL_N4"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
            
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.BackColor = basModule.WhiteColor
                sprLsn.BackColorStyle = BackColorStyleUnderGrid
            sprLsn.BlockMode = False

            sprLsn.ColsFrozen = 5

        '>> spread lock
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.Lock = True
                sprLsn.Protect = True
            sprLsn.BlockMode = False
            
            
            sprLsn.Row = 1:         sprLsn.Row2 = 1
            sprLsn.Col = 3:         sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.BackColor = &HC0C0FF
                sprLsn.BackColorStyle = BackColorStyleUnderGrid
            sprLsn.BlockMode = False
            
            sprLsn.Row = 1:         sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 3:         sprLsn.Col2 = 3
            sprLsn.BlockMode = True
                sprLsn.BackColor = &HC0C0FF
                sprLsn.BackColorStyle = BackColorStyleUnderGrid
            sprLsn.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Find_Gwamok_to_STD_TOT = nRet
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ���� ������û���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� ������û���� ��ȸ"
    
    Find_Gwamok_to_STD_TOT = nRet
End Function






Private Sub sprData_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprData
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
    
End Sub


Private Sub sprBan_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprBan
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 5
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 7:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag)
            .Value = 0
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = 5
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 7:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row
            .Value = 1
        
        .Tag = Trim(CStr(Row))
        
    End With
    
End Sub







'## �л��ο� ����
Private Sub sprLsn_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nChk        As Long
    
    Dim nColor      As Long
    
    Dim ni          As Long
    Dim nj          As Long
    
    Dim nChkTot     As Integer
    
    lblStatus.Caption = ""
    
    If Row <= 1 Then
        lblStatus.Caption = "������ �����ϼ���."
        Exit Sub
    End If
    If Col < 4 Then
        lblStatus.Caption = "������ �����ϼ���."
        Exit Sub
    End If
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"
            If Col > 15 Then
                lblStatus.Caption = "������ �����ϼ���."
                Exit Sub
            End If
        Case "02"
            If Col > 12 Then
                lblStatus.Caption = "������ �����ϼ���."
                Exit Sub
            End If
    End Select
    
    
    With sprLsn
        .Row = Row
        .Col = Col
        
        If Trim(.Text) = "" Then
            lblStatus.Caption = "�ش���� ���ð��� �л��ο��� �����ϴ�."
            Exit Sub
        End If
        If Trim(.Text) = 0 Then
            lblStatus.Caption = "�ش���� ���ð��� �л��ο��� �����ϴ�."
            Exit Sub
        End If
        If .BackColor <> basModule.WhiteColor Then
            lblStatus.Caption = "�̹� ���õǾ��� �����Դϴ�."
            Exit Sub
        End If
        
        
        nChk = 0
        nColor = 0
        For ni = 1 To sprBan.MaxRows Step 1
            sprBan.Row = ni
            sprBan.Col = sprBan.MaxCols
            
            If sprBan.Value = 1 Then
                If nChk > 0 Then
                    lblStatus.Caption = "�����׸��� 2�� �̻��Դϴ�."
                    
                    sprBan.Value = 0
                    Exit Sub
                End If
                
                sprBan.Col = 6
                
                If sprBan.BackColor = &HFFFFFF Then
                    lblStatus.Caption = "�� ���� ���� ����Դϴ�. ���� ����ϼ���."
                    
                    Exit Sub
                End If
                
                nColor = sprBan.BackColor       '< color
                If nChk = 0 Then nChk = sprBan.Col
                
            End If
        Next ni
        
        If nChk = 0 Then            '< �� ������ ���� �����.
            lblStatus.Caption = "������ ���� �����ϼ���."
            Exit Sub
        End If
        
        For nRow = 2 To .MaxRows Step 1
        '> ���콺 ������ ������ �ƴϾ�� �Ѵ�.
            If nRow <> Row Then
                .Row = nRow
                .Col = Col
                    
                If nColor = .BackColor Then
                    lblStatus.Caption = "�̹� ���õǾ��� �����Դϴ�."
                    Exit Sub
                End If
            End If
        Next nRow
        
        
        nChkTot = 0
        For nRow = 2 To .MaxRows Step 1
            For nCol = 4 To 15 Step 1
                .Row = nRow
                .Col = nCol
                
                If .BackColor = nColor Then
                    nChkTot = nChkTot + 1
                End If
            Next nCol
        Next nRow
        
        If nChkTot > 3 Then
            lblStatus.Caption = "�ش���� ������ 4���̻� �����Ͽ����ϴ�."
            Exit Sub
        End If
        
        .Row = Row:     .Row2 = Row
        .Col = Col:     .Col2 = Col
        .BlockMode = True
            .BackColor = nColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
    End With
End Sub

Private Sub sprLsn_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    lblStatus.Caption = ""
    
    If Row <= 1 Then
        lblStatus.Caption = "������ �����ϼ���."
        Exit Sub
    End If
    If Col < 4 Then
        lblStatus.Caption = "������ �����ϼ���."
        Exit Sub
    End If
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"
            If Col > 15 Then
                lblStatus.Caption = "������ �����ϼ���."
                Exit Sub
            End If
        Case "02"
            If Col > 12 Then
                lblStatus.Caption = "������ �����ϼ���."
                Exit Sub
            End If
    End Select
    
    With sprLsn
        .Row = Row:     .Row2 = .Row
        .Col = Col:     .Col2 = .Col
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With
    
End Sub


























