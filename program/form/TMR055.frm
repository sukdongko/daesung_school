VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMR055 
   Caption         =   "�ð�ǥ ����� >> ��ü�ð�ǥ ���� - ���纰"
   ClientHeight    =   10815
   ClientLeft      =   75
   ClientTop       =   1980
   ClientWidth     =   19095
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   19095
   WindowState     =   2  '�ִ�ȭ
   Begin VB.Frame Frame5 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame5"
      Height          =   6195
      Left            =   60
      TabIndex        =   9
      Top             =   6090
      Width           =   18945
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame4"
         Height          =   6135
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   18885
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "�ð�ǥ ũ�Ժ���"
            Height          =   210
            Index           =   0
            Left            =   1740
            TabIndex        =   17
            Top             =   330
            Width           =   1905
         End
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "�ð�ǥ �۰Ժ���"
            Height          =   210
            Index           =   1
            Left            =   1740
            TabIndex        =   16
            Top             =   60
            Width           =   1905
         End
         Begin VB.CommandButton cmdDelTimeTable 
            Caption         =   "�ð�ǥ ���� ����"
            Height          =   500
            Left            =   6060
            TabIndex        =   12
            Top             =   30
            Width           =   2595
         End
         Begin VB.CommandButton cmdShowTimeTable 
            Caption         =   "��ü�ð�ǥ ��ȸ"
            Height          =   500
            Left            =   3660
            TabIndex        =   11
            Top             =   30
            Width           =   1905
         End
         Begin FPSpread.vaSpread sprTimeTable 
            Height          =   5535
            Left            =   0
            TabIndex        =   13
            Top             =   570
            Width           =   18855
            _Version        =   393216
            _ExtentX        =   33258
            _ExtentY        =   9763
            _StockProps     =   64
            EditEnterAction =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "TMR055.frx":0000
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "��ü �ð�ǥ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E5A75&
            Height          =   210
            Left            =   150
            TabIndex        =   15
            Top             =   150
            Width           =   3075
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame3"
      Height          =   5985
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   19005
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   5925
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   18945
         Begin VB.CommandButton cmdFind_TeacherData 
            Caption         =   "���纰 ���� ��ȸ"
            Height          =   495
            Left            =   3780
            TabIndex        =   5
            Top             =   90
            Width           =   1845
         End
         Begin VB.CommandButton cmdWorkTableSave 
            Caption         =   "��ü �ð�ǥ�� �ݿ��ϱ� (�ð�ǥ ����)"
            Height          =   495
            Left            =   6120
            TabIndex        =   6
            Top             =   90
            Width           =   3945
         End
         Begin FPSpread.vaSpread sprWork 
            Height          =   5235
            Left            =   30
            TabIndex        =   7
            Top             =   660
            Width           =   18885
            _Version        =   393216
            _ExtentX        =   33311
            _ExtentY        =   9234
            _StockProps     =   64
            EditEnterAction =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "TMR055.frx":446A
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   18000
            Top             =   360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "��� ������ �ݺ� ���ð��� �ü������� Ŭ�� ��  S �� �����ø� ���� �Էµ˴ϴ�."
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   2
            Left            =   10530
            TabIndex        =   18
            Top             =   390
            Width           =   7035
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "�۾� �ð�ǥ ���̺�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E5A75&
            Height          =   210
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   3075
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  '����
            Caption         =   "lblStatus"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   10560
            TabIndex        =   8
            Top             =   150
            Width           =   8055
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1485
      Left            =   9000
      TabIndex        =   0
      Top             =   12930
      Visible         =   0   'False
      Width           =   3165
      Begin VB.CommandButton cmdHeaderReview 
         Caption         =   "WorkSpreadHeader"
         Height          =   555
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   2445
      End
      Begin VB.CommandButton cmdTimeTableHeaderReview 
         Caption         =   "TimetableSpreadHeader"
         Height          =   555
         Left            =   120
         TabIndex        =   1
         Top             =   780
         Width           =   2445
      End
   End
End
Attribute VB_Name = "TMR055"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM055
'   �� ��  �� �� : ��ü�ð�ǥ ���� - ���纰
'
'   ��   ��   �� : 2007/11/28
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################


Option Explicit


Private Const nRowHeight = 14

Private nTtRowHeight            As Long
Private nTtColWidth             As Long

Private nViewCnt                As Integer

Private Type tWorkTimeTable
    ACID        As String
    LSNCD       As String
    LESSON      As String
    WEEK        As String
    SISUCD      As String
    SISU        As String
    TRX_CL      As String
End Type
Private uWorkTimeTable() As tWorkTimeTable

Private Sub Form_Load()
    
    With sprWork
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
        .MaxCols = 0
        
        .RowHeaderCols = 1
        .ColHeaderRows = 1
        
        
    End With
    
    With sprTimeTable
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
        .MaxCols = 0
        
        .RowHeaderCols = 1
        .ColHeaderRows = 1
        
    End With
    
    Me.Tag = "LOAD"
    
    
        optView(0).Value = False
        optView(1).Value = True

'        If optView(0).Value = True Then
'            nTtRowHeight = 25
'            nTtColWidth = 6
'        ElseIf optView(1).Value = True Then
'            nTtRowHeight = 15
'            nTtColWidth = 5
'        End If

        nViewCnt = 0
        lblStatus.Caption = ""

        cmdFind_TeacherData.Tag = ""
        cmdShowTimeTable.Tag = ""
        
    Me.Tag = ""
    
End Sub


Private Sub Form_Activate()
    
    If nViewCnt = 0 Then
        
        With sprWork
            .RowHeaderCols = 1
            .ColHeaderRows = 3
            
            .MaxCols = 10 + 70
            
            Call Make_WorkSpread_Header
        End With
        
        With sprTimeTable
            .RowHeaderCols = 4
            .ColHeaderRows = 3
            
            .MaxCols = 70
            
            Call Make_TimeTableSpread_Header
        End With
        
    End If
    
    nViewCnt = 1
    
End Sub




'## sprWork �� spread ��� �����
Private Sub cmdHeaderReview_Click()
    Call Make_WorkSpread_Header
End Sub

Private Sub Make_WorkSpread_Header()

    Dim sRet        As String
    Dim nCols       As Long
    
    Dim nTmp        As Long
    
    With sprWork
        
        If optView(0).Value = True Then
            nTtRowHeight = 16
            nTtColWidth = 3
        ElseIf optView(1).Value = True Then
            nTtRowHeight = 16
            nTtColWidth = 3
        End If
        
        .Row = SpreadHeader
        
        .Col = SpreadHeader:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 1:       .Text = "�ü��ڵ�":         .ColWidth(.Col) = 8:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 2:       .Text = "����":             .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 3:       .Text = "�� �ü�":          .ColWidth(.Col) = 4:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 4:       .Text = "����":             .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 5:       .Text = "��":               .ColWidth(.Col) = 5:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 6:       .Text = "���ɽü�":         .ColWidth(.Col) = 5:        .AddCellSpan .Col, .Row, 1, 3
        
        '<< �� ���� combo box >>
        .Col = 7:       .Text = "��":               .ColWidth(.Col) = 11:       .AddCellSpan .Col, .Row, 1, 3
        .Col = 8:       .Text = "����":             .ColWidth(.Col) = 4:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 9:       .Text = " ":                .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 10:      .Text = " ":                .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        
        
        '<< ���� ����� >>
        For nCols = 1 To 7 Step 1
            Select Case nCols
                Case 1
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "2"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 2
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "ȭ"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "3"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 3
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "4"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 4
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "5"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 5
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "6"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 6
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "7"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 7
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "1"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
            End Select
        Next nCols
        
        .Row = SpreadHeader + 1:    .RowHidden = True
        
    End With
End Sub


Private Sub cmdFind_TeacherData_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nWorkRow    As Long
    Dim nWorkCol    As Long
    
    Dim sRet        As String
    Dim sDivN()     As String
    Dim sDivT()     As String
    Dim sLsn        As String
    
    Dim nChaSisu    As Long
    
    
    Dim sSisuCD     As String
    
    
    On Error GoTo ErrStmt
    
    Call Make_WorkSpread_Header
    
    sStr = ""
    sStr = sStr & "  SELECT A.SISUCD, A.TCRNM, A.SISU, A.SUBJNM, A.TCR_CL,"
    sStr = sStr & "         NVL(A.SISU,0)-NVL(B.SEL_SISU,0) AS SEL_SISU"
    sStr = sStr & "    FROM (SELECT A.ACID, A.SISUCD, A.TCRNM, NVL(B.SISU,0) AS SISU, A.SUBJNM, A.TCR_CL "
    sStr = sStr & "            FROM SDTCR01TB A,"
    sStr = sStr & "                 (SELECT ACID, SISUCD, SUM(SISU) AS SISU"
    sStr = sStr & "                    FROM SDTCR11TB"
    sStr = sStr & "                   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                   GROUP BY ACID, SISUCD"
    sStr = sStr & "                  ) B"
    sStr = sStr & "           WHERE A.ACID   = B.ACID (+)"
    sStr = sStr & "             AND A.SISUCD = B.SISUCD (+)"
    sStr = sStr & "             AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "           ORDER BY A.TCRNM"
    sStr = sStr & "          ) A,"
    sStr = sStr & "         (SELECT ACID, SISUCD, SUM(SISU) AS SEL_SISU"
    sStr = sStr & "            FROM SDTRX50TB"
    sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "           GROUP BY ACID, SISUCD"
    sStr = sStr & "          ) B"
    sStr = sStr & "   WHERE A.ACID   = B.ACID (+)"
    sStr = sStr & "     AND A.SISUCD = B.SISUCD (+)"
    sStr = sStr & "   ORDER BY A.TCRNM, A.SUBJNM"
    
    
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
    
    'XXX
    
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        sprWork.MaxRows = 0
        
        

        '>> ������ �ֱ� --------------------------------------------------------------------
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                
                sprWork.MaxRows = sprWork.MaxRows + 1
                sprWork.Row = sprWork.MaxRows:      sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                
                sprWork.Col = 1:                    sTmp = ""
                    sSisuCD = ""
                    If IsNull(.Fields("SISUCD")) = False Then
                        sTmp = Trim(.Fields("SISUCD")):     sSisuCD = sTmp      '< �ü��ڵ�
                    End If
                    Call basFunction.Set_SprType_Text(sprWork, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprWork.Col = sprWork.Col + 1:      sTmp = ""
                    If IsNull(.Fields("TCRNM")) = False Then
                        sTmp = Trim(.Fields("TCRNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprWork, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprWork.Col = sprWork.Col + 1:      nTmp = 0
                    If IsNumeric(.Fields("SISU")) = True Then
                        nTmp = CDbl(.Fields("SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprWork, 0, -99999, 99999, "", nTmp)
                    If nTmp <= 0 Then
                        sprWork.ForeColor = &HFF&
                    Else
                        sprWork.ForeColor = &H80000008
                    End If
                    nChaSisu = nTmp
                    
                sprWork.Col = sprWork.Col + 1:      sTmp = ""
                    If IsNull(.Fields("SUBJNM")) = False Then
                        sTmp = Trim(.Fields("SUBJNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprWork, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                sprWork.Col = sprWork.Col + 1:      sTmp = ""
                    If IsNull(.Fields("TCR_CL")) = False Then
                        sprWork.Row2 = sprWork.Row
                        sprWork.Col2 = sprWork.Col
                        sprWork.BlockMode = True
                            sprWork.BackColor = CLng(.Fields("TCR_CL"))
                            sprWork.BackColorStyle = BackColorStyleUnderGrid
                        sprWork.BlockMode = False
                    Else
                        sprWork.Row2 = sprWork.Row
                        sprWork.Col2 = sprWork.Col
                        sprWork.BlockMode = True
                            sprWork.BackColor = basModule.WhiteColor
                            sprWork.BackColorStyle = BackColorStyleUnderGrid
                        sprWork.BlockMode = False
                    End If
                    
                sprWork.Col = sprWork.Col + 1:      nTmp = 0
                    If IsNumeric(.Fields("SEL_SISU")) = True Then
                        nTmp = CDbl(.Fields("SEL_SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprWork, 0, -99999, 99999, "", nTmp)
                    nChaSisu = nChaSisu - nTmp
                    If nTmp > 0 Then
                        sprWork.ForeColor = &HFF0000
                    Else
                        sprWork.ForeColor = &HFF&
                    End If
                
                '<< �ݳ��� ��ȸ >>
                sprWork.Col = sprWork.Col + 1:      sRet = ""
                    
                    '  -- test --
                        sRet = "<<�� ����>>[T]ALL[N]"
                        sRet = sRet & Get_SisuCD_to_Lsn(sSisuCD)
                    
'                        sRet = "�ι�1" & "[T]" & "00001" & "[N]" & _
'                               "�ι�2" & "[T]" & "00002" & "[N]" & _
'                               "�ι�3" & "[T]" & "00003" & "[N]"
                                                  
                    If sRet > " " Then
                        sLsn = ""
                        
                        sDivN = Split(sRet, "[N]", -1, vbTextCompare)
                        If UBound(sDivN) > 0 Then
                        
                            For ni = 0 To UBound(sDivN) - 1 Step 1
                                
                                sDivT = Split(sDivN(ni), "[T]", -1, vbTextCompare)
                                
                                If UBound(sDivT) = 1 Then
                                    If ni > 0 Then sLsn = sLsn & Chr$(9)
                                    sLsn = sLsn & sDivT(0) & Space(30) & sDivT(1)
                                    
                                End If
                                
                            Next ni
                        
                            sprWork.CellType = CellTypeComboBox
                            sprWork.TypeComboBoxClear 1, 1
                            
                            sprWork.TypeComboBoxList = sLsn
                            sprWork.TypeComboBoxEditable = False
                            sprWork.TypeComboBoxMaxDrop = 5
                            sprWork.TypeComboBoxCurSel = 0
                            'sprWork.TypeComboBoxWidth = 1
                            
                        End If
                    End If
                    
                
                '����
                sprWork.Col = sprWork.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprWork):       sprWork.Value = 0
                
                '����
                sprWork.Col = sprWork.Col + 1
                
                '<< ���ϳ��� >>
                
                ' CLEAR
                    For nWorkRow = 1 To sprWork.MaxRows Step 1
                        sprWork.Row = nWorkRow
                        For nWorkCol = 11 To sprWork.MaxCols Step 1
                            sprWork.Col = nWorkCol
                                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                        Next nWorkCol
                    Next nWorkRow
                    sprWork.Row = 1:   sprWork.Row2 = sprWork.MaxRows
                    sprWork.Col = 11:  sprWork.Col2 = sprWork.MaxCols
                    sprWork.BlockMode = True
                        sprWork.BackColor = basModule.WhiteColor
                        sprWork.BackColorStyle = BackColorStyleUnderGrid
                    sprWork.BlockMode = False
                
                
                .MoveNext
                
            Next nRec
        End If
        
        '>> �ʱ�ȭ   -----------------------------------------------------------------------
        sprWork.SetCellBorder 6, 1, 6, sprWork.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        sprWork.SetCellBorder 7, 1, 7, sprWork.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        
        For nWorkRow = 1 To sprWork.MaxRows Step 1
            sprWork.Row = nWorkRow
            For nWorkCol = 10 To sprWork.MaxCols Step 1            '< ���� �������� CLEAR
                sprWork.Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                
                If nWorkCol Mod 10 = 0 Then
                    sprWork.SetCellBorder sprWork.Col, sprWork.Row, sprWork.Col, sprWork.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                End If
            Next nWorkCol
        Next nWorkRow
        
        sprWork.Row = 1:   sprWork.Row2 = sprWork.MaxRows
        sprWork.Col = 1:   sprWork.Col2 = 6
        sprWork.BlockMode = True
            sprWork.Lock = True
            sprWork.Protect = True
        sprWork.BlockMode = False
        
        sprWork.Row = 1:   sprWork.Row2 = sprWork.MaxRows
        sprWork.Col = 8:   sprWork.Col2 = 10
        sprWork.BlockMode = True
            sprWork.Lock = True
            sprWork.Protect = True
        sprWork.BlockMode = False
        
        
        sprWork.Row = 1:   sprWork.Row2 = sprWork.MaxRows
        sprWork.Col = 1:   sprWork.Col2 = 4
        sprWork.BlockMode = True
            sprWork.BackColor = basModule.WhiteColor
            sprWork.BackColorStyle = BackColorStyleUnderGrid
        sprWork.BlockMode = False
        
        sprWork.Row = 1:   sprWork.Row2 = sprWork.MaxRows
        sprWork.Col = 6:   sprWork.Col2 = sprWork.MaxCols
        sprWork.BlockMode = True
            sprWork.BackColor = basModule.WhiteColor
            sprWork.BackColorStyle = BackColorStyleUnderGrid
        sprWork.BlockMode = False
        
        
        
        
    End With
    
    If cmdFind_TeacherData.Tag <> "SAVE" Then
        MsgBox "�� ���纰 ���� �����ϼ���." & vbCrLf & _
               "�� ���ý� ��� ������ �ð�ǥ ������ �� �� �ֽ��ϴ�.", vbInformation + vbOKOnly, "����ü����� ��ȸ"
               
        cmdFind_TeacherData.Tag = ""
    End If
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    MsgBox "���纰 �ѽü� ���� ��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "����ü����� ��ȸ"
    
    On Error GoTo 0
End Sub

Private Function Get_SisuCD_to_Lsn(ByVal aSisuCD As String) As String

    Dim sRet        As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    sRet = ""
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT TRIM(B.LSNNM)||'[T]'||TRIM(A.LSNCD)||'[N]' AS LSN"
    sStr = sStr & "    FROM SDTCR11TB A, SDLSN01TB B"
    sStr = sStr & "   Where A.ACID = B.ACID"
    sStr = sStr & "     AND A.LSNCD = B.LSNCD"
    sStr = sStr & "     AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.SISUCD= " & aSisuCD
    
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
        sTmp = ""
        If .RecordCount > 0 Then
            .MoveFirst
            For nRec = 1 To .RecordCount Step 1
                If IsNull("LSN") = False Then
                    sTmp = sTmp & Trim(.Fields("LSN"))
                End If
                
                .MoveNext
            Next nRec
        End If
    End With
    
    If sTmp > " " Then
        sRet = sTmp
    End If
    
    Get_SisuCD_to_Lsn = sRet
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "�� ��ȸ�� �����Դϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�� ��ȸ"
    
    Get_SisuCD_to_Lsn = sRet
    On Error GoTo 0
End Function




'>> �� ���
Private Sub sprWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nColor      As Long
    
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nEXE        As Long
    Dim ni          As Long
    
    Dim sSchCD      As String
    Dim sSisuCD     As String
    
    On Error GoTo CancelColor
    
    If Col = 5 And Row >= 1 Then
        With dlgCommon
            .CancelError = True
            .ShowColor
            
            nColor = .color
            
            '## ��ҽÿ� CancelColor �� �Ѿ��.
        End With
        
        On Error GoTo 0
        On Error GoTo ErrStmt
        
        
        sSchCD = Trim(basModule.SchCD)                                                      '< �п�
        sprWork.Row = Row:      sprWork.Col = 1:        sSisuCD = Trim(sprWork.Text)        '< �ü��ڵ�
        
        
        basDataBase.DBConn.BeginTrans

        Set DBCmd = New ADODB.Command
        Set DBParam = New ADODB.Parameter
    
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        
        sStr = ""
        sStr = sStr & "  UPDATE SDTCR01TB"
        sStr = sStr & "     SET TCR_CL =  " & Trim(CStr(nColor))
        sStr = sStr & "   WHERE ACID   = '" & sSchCD & "'"
        sStr = sStr & "     AND SISUCD =  " & sSisuCD
        
        
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
    '    '>> color
    '        nTmp = aColor
    '            Set DBParam = DBCmd.CreateParameter("TRX_CL", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '    '>> �п�
    '        sTmp = sSchCD
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> ssisucd
    '        nTmp = CLng(sSisuCD)
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
                            
            sprWork.Row2 = sprWork.Row
            sprWork.Col = Col:      sprWork.Col2 = sprWork.Col
            sprWork.BlockMode = True
                sprWork.BackColor = nColor
                sprWork.BackColorStyle = BackColorStyleUnderGrid
            sprWork.BlockMode = False
            
            MsgBox "������ ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���� �����ϱ�"
        Else
            basDataBase.DBConn.RollbackTrans
            
            sprWork.Row2 = sprWork.Row:
            sprWork.Col = Col:      sprWork.Col2 = sprWork.Col
            sprWork.BlockMode = True
                sprWork.BackColor = basModule.WhiteColor
                sprWork.BackColorStyle = BackColorStyleUnderGrid
            sprWork.BlockMode = False
            
            MsgBox "���� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �����ϱ�"
            
        End If
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
CancelColor:
    MsgBox "��������Ͽ����ϴ�.", vbExclamation + vbOKOnly, "���� �����ϱ�"
    Exit Sub
    
ErrStmt:
    MsgBox "���� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �����ϱ�"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub


























Private Sub cmdTimeTableHeaderReview_Click()
    Call Make_TimeTableSpread_Header
End Sub

Private Sub Make_TimeTableSpread_Header()

    Dim sRet        As String
    Dim nCols       As Long
    
    Dim nTmp        As Long
    Dim nRow        As Long
    
    With sprTimeTable
        
        If optView(0).Value = True Then
            nTtRowHeight = 19
            nTtColWidth = 6
        ElseIf optView(1).Value = True Then
            nTtRowHeight = 11
            nTtColWidth = 4
        End If
        
        .Row = SpreadHeader
        
        .Col = SpreadHeader:        .Text = "����":         .ColWidth(.Col) = 6:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 1:    .Text = "����":         .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 2:    .Text = "��  �ü�":     .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 3:    .Text = "����  �ü�":   .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        
        
        '<< ���� ����� >>
        For nCols = 1 To 7 Step 1
            Select Case nCols
                Case 1
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "2"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 2
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "ȭ"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "3"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 3
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "4"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 4
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "5"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 5
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "6"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 6
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "7"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 7
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "��"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column�� ������ ���¿��� ó��
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "1"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
            End Select
        Next nCols
        
        .Row = SpreadHeader + 1:        .RowHidden = True
        
        If .MaxRows > 0 Then
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow:        .RowHeight(.Row) = nTtRowHeight
            Next nRow
        End If
        
    End With
End Sub

' ��ü �ð�ǥ ��ȸ <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub cmdShowTimeTable_Click()
    
    Dim nLsnCount       As Long
    
    Dim DBCmd           As ADODB.Command
    Dim DBRec           As ADODB.Recordset
    Dim DBParam         As ADODB.Parameter
    
    Dim nLength         As Long
    Dim sStr            As String
    Dim ni              As Integer
    Dim nRec            As Long
    Dim sTmp            As String
    Dim nTmp            As Long
    
    Dim nRow            As Long
    Dim nCol            As Long
    Dim nTcrRow         As Long
    
    Dim sTcrNM          As String
    Dim sLesson         As String
    Dim sWeeks          As String
    Dim sTcr_CL         As String
    Dim sDisp_Text      As String
    
    sprTimeTable.MaxRows = 0
    Call Make_TimeTableSpread_Header
    On Error GoTo ErrStmt
    
    
    '## ��ü���� ��� ��ȸ
    sStr = ""
    sStr = sStr & "  SELECT ACID, TCRNM, SISUCD, SUBJNM,"
    sStr = sStr & "         DAMIM, DISP, LESSON , WEEKS,"
    sStr = sStr & "         GET_TCR_T_SISU(ACID, TCRNM) AS T_SISU,"
    sStr = sStr & "         GET_TCR_S_SISU(ACID, TCRNM) AS S_SISU,"
    sStr = sStr & "         TCR_CL"
    sStr = sStr & "    FROM (SELECT A.ACID, A.TCRNM,"
    sStr = sStr & "                 A.SISUCD, A.SUBJNM,"
    sStr = sStr & "                 DAMIM, DISP, LESSON , WEEKS, TRX_CL AS TCR_CL"
    sStr = sStr & "            FROM (SELECT ACID, SISUCD, TCRNM, SUBJNM"
    sStr = sStr & "                    FROM SDTCR01TB"
    sStr = sStr & "                   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                  ) A,"
    sStr = sStr & "                 (SELECT ACID,"
    sStr = sStr & "                         SISUCD,"
    sStr = sStr & "                         GET_DAMIM_TCR01(ACID, LSNCD) AS DAMIM,"
    sStr = sStr & "                         GET_KEAYOL_N_LSN_TCR01(ACID, LSNCD) AS DISP,"
    sStr = sStr & "                         LESSON,"
    sStr = sStr & "                         WEEKS,"
    sStr = sStr & "                         TRX_CL"
    sStr = sStr & "                    From SDTRX50TB"
    sStr = sStr & "                   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                  ) B"
    sStr = sStr & "           WHERE A.ACID   = B.ACID (+)"
    sStr = sStr & "             AND A.SISUCD = B.SISUCD (+)"
    sStr = sStr & "          )"
    sStr = sStr & "   ORDER BY ACID, TCRNM"
    
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
                
''>> �п�
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
                
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    
    If DBRec.RecordCount > 0 Then
    
        DBRec.MoveFirst
        For nRec = 1 To DBRec.RecordCount Step 1
                        
            '> ��������� ó��
            nTcrRow = 0
            If IsNull(DBRec.Fields("TCRNM")) = False Then
                sTcrNM = Trim(DBRec.Fields("TCRNM"))
                
                For nRow = 1 To sprTimeTable.MaxRows Step 1     '< ���� ����� ��ȸ
                    sprTimeTable.Row = nRow
                    sprTimeTable.Col = SpreadHeader
                    If StrComp(Trim(sprTimeTable.Text), sTcrNM, vbTextCompare) = 0 Then
                        nTcrRow = nRow                          '< ���簡 ��ġ�� row
                        
                        sprTimeTable.Row = nRow
                        sprTimeTable.Col = SpreadHeader:        sprTimeTable.Text = Trim(DBRec.Fields("TCRNM"))
                                                   
                        If IsNull(DBRec.Fields("DAMIM")) = False Then
                            sprTimeTable.Col = SpreadHeader + 1
                            sprTimeTable.Text = Trim(DBRec.Fields("DAMIM"))
                        End If
                                                   
                        If IsNumeric(DBRec.Fields("T_SISU")) = True Then
                            sprTimeTable.Col = SpreadHeader + 2
                            If IsNumeric(sprTimeTable.Text) = True Then
                                sprTimeTable.Col = SpreadHeader + 2
                                    nTmp = CLng(DBRec.Fields("T_SISU"))
                                    sprTimeTable.Text = Trim(CStr(nTmp))
                            Else
                                sprTimeTable.Col = SpreadHeader + 2
                                sprTimeTable.Text = Trim(CStr(DBRec.Fields("T_SISU")))
                            End If
                            
                        End If
                        If IsNumeric(DBRec.Fields("S_SISU")) = True Then
                            sprTimeTable.Col = SpreadHeader + 3
                            If IsNumeric(sprTimeTable.Text) = True Then
                                sprTimeTable.Col = SpreadHeader + 3
                                    nTmp = CLng(DBRec.Fields("S_SISU"))
                                    sprTimeTable.Text = Trim(CStr(nTmp))
                            Else
                                sprTimeTable.Col = SpreadHeader + 3
                                sprTimeTable.Text = Trim(CStr(DBRec.Fields("S_SISU")))
                            End If
                            
                        End If
                        
                        Exit For
                    End If
                Next nRow
                
                If nTcrRow = 0 Then       '>> ���系�� �߰�
                    sprTimeTable.MaxRows = sprTimeTable.MaxRows + 1
                    sprTimeTable.Row = sprTimeTable.MaxRows:        sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                    
                    
                    nTcrRow = sprTimeTable.Row                  '< ���ο� row�߰��� ����
                    
                    sprTimeTable.Col = SpreadHeader
                        sprTimeTable.Text = Trim(DBRec.Fields("TCRNM"))
                    
                    sprTimeTable.Col = SpreadHeader + 1
                        sprTimeTable.Text = ""                  '< �ʱ�ȭ
                    sprTimeTable.Col = SpreadHeader + 2
                        sprTimeTable.Text = ""                  '< �ʱ�ȭ
                    sprTimeTable.Col = SpreadHeader + 3
                        sprTimeTable.Text = ""                  '< �ʱ�ȭ
                        
                        
                    If IsNull(DBRec.Fields("DAMIM")) = False Then
                        sprTimeTable.Col = SpreadHeader + 1
                        sprTimeTable.Text = Trim(DBRec.Fields("DAMIM"))
                    End If
                    
                    
                    If IsNumeric(DBRec.Fields("T_SISU")) = True Then
                        sprTimeTable.Col = SpreadHeader + 2
                        If IsNumeric(sprTimeTable.Text) = True Then
                            sprTimeTable.Col = SpreadHeader + 2
                                nTmp = CLng(DBRec.Fields("T_SISU"))
                                sprTimeTable.Text = Trim(CStr(nTmp))
                        Else
                            sprTimeTable.Col = SpreadHeader + 2
                            sprTimeTable.Text = Trim(CStr(DBRec.Fields("T_SISU")))
                        End If
                    Else
                        sprTimeTable.Col = SpreadHeader + 2
                        If Trim(sprTimeTable.Text) = "" Then
                            sprTimeTable.Text = " "
                        Else
                            sprTimeTable.Text = Trim(CStr(DBRec.Fields("T_SISU")))
                        End If
                        
                    End If
                    If IsNumeric(DBRec.Fields("S_SISU")) = True Then
                        sprTimeTable.Col = SpreadHeader + 3
                        If IsNumeric(sprTimeTable.Text) = True Then
                            sprTimeTable.Col = SpreadHeader + 3
                                nTmp = CLng(DBRec.Fields("S_SISU"))
                                sprTimeTable.Text = Trim(CStr(nTmp))
                        Else
                            sprTimeTable.Col = SpreadHeader + 3
                            sprTimeTable.Text = Trim(CStr(DBRec.Fields("S_SISU")))
                        End If
                    Else
                        sprTimeTable.Col = SpreadHeader + 3
                        If Trim(sprTimeTable.Text) = "" Then
                            sprTimeTable.Text = " "
                        Else
                            sprTimeTable.Text = Trim(CStr(DBRec.Fields("T_SISU")))
                        End If
                    End If
                    
                    
                End If
                
                
            '<< ������ ���� ==================================================================================
                
                ' nTcrRow <- ���� ��ġ�� row
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    sTcr_CL = "":       If IsNull(DBRec.Fields("TCR_CL")) = False Then sTcr_CL = Trim(DBRec.Fields("TCR_CL"))                   ' COLOR
                    sDisp_Text = "":    If IsNull(DBRec.Fields("DISP")) = False Then sDisp_Text = Trim(DBRec.Fields("DISP"))                    ' ������ ������
                                        If IsNull(DBRec.Fields("SUBJNM")) = False Then sDisp_Text = sDisp_Text & vbCrLf & Trim(DBRec.Fields("SUBJNM"))
                    
                    
                    sprTimeTable.Row = nTcrRow              '< ó���� row
                    
                        Select Case sWeeks
                            Case "2"
                                sprTimeTable.Col = 1 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "3"
                                sprTimeTable.Col = 11 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "4"
                                sprTimeTable.Col = 21 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "5"
                                sprTimeTable.Col = 31 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "6"
                                sprTimeTable.Col = 41 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "7"
                                sprTimeTable.Col = 51 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                            Case "1"
                                sprTimeTable.Col = 61 + CLng(sLesson) - 1

                                '< setting rows and col & display data  >
                                Call basFunction.Set_SprType_Text(sprTimeTable, "center", "center", basFunction.LenKor(sDisp_Text), sDisp_Text)
                                    sprTimeTable.TypeEditMultiLine = True
                                If sTcr_CL > " " Then
                                    sprTimeTable.Row2 = sprTimeTable.Row
                                    sprTimeTable.Col2 = sprTimeTable.Col
                                    sprTimeTable.BlockMode = True
                                        sprTimeTable.BackColor = CLng(sTcr_CL)
                                    sprTimeTable.BlockMode = False
                                End If

                        End Select
                End If
            End If      ' �����
            
            DBRec.MoveNext
        Next nRec
        
    End If      ' recordcount
                        
       
    With sprTimeTable
        For nCol = 10 To .MaxCols Step 1
            .Col = nCol
            If nCol Mod 10 = 0 Then
                .SetCellBorder .Col, 1, .Col, .MaxCols, 2, basModule.SectionColor1, CellBorderStyleSolid
            End If
        Next nCol
                
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
                
    End With
    
    If cmdShowTimeTable.Tag <> "SAVE" Then
        MsgBox "��ü�ð�ǥ ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ü �ð�ǥ ��ȸ"
        
        cmdShowTimeTable.Tag = ""
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "��ü�ð�ǥ ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "��ü �ð�ǥ ��ȸ"
    
End Sub





'## �� ó��
Private Sub sprWork_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sSisuCD     As String
    Dim sGGbn       As String
    Dim sTcr_CL     As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprWork
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 6:           .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 6:       .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        
        
        '>> 1�� ���õ� �κ��� ���尡�ɻ��·� �ٲپ� ��.
        If Col >= 11 Then
            .Row = Row
            .Col = Col
            
            If .Text = "1" Then
                .Text = "S"
                
                .SetCellBorder .Col, .Row, .Col, .Row, 16, basModule.SectionColor1, CellBorderStyleSolid
                
                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = &HC0C0C0
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = 8:       .Value = 1
                
            ElseIf .Text = "S" Then
                .Col = Col:     .Text = "1"
                .SetCellBorder .Col, .Row, .Col, .Row, 16, basModule.GridColor2, CellBorderStyleSolid
                
                .Row = Row
                .Col = 1:       sSisuCD = Trim(.Text)
                    Call Get_Lsn_Detail_Note(sSisuCD, sGGbn, sTcr_CL)     '< �� ����[������.Ž��]/ �� ��
                
                .Row2 = .Row
                .Col = Col:     .Col2 = Col
                .BlockMode = True
                    .BackColor = sTcr_CL
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
            End If
        End If
    End With
    
End Sub



Private Sub sprWork_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim nTcrRow     As Long
    Dim sSchCD      As String
    Dim sSisuCD     As String
    Dim sGGbn     As String       ' ��������
    Dim sTcr_CL     As String
    Dim sTeacher    As String
    Dim sGwamok     As String
    Dim sLsnCD      As String
    
    Dim nWTotSisu   As Long
    Dim nWLsnSisu   As Long
    
    Dim nWorkRow    As Long
    Dim nWorkCol    As Long
    
    nTcrRow = Row       '< �۾���� row
    
    With sprWork
    
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 6:           .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 6:       .Col2 = 10
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
            
            
            
        If Col = 7 Then
            .Row = Row
            .Col = Col
            
            sLsnCD = Trim(Right(.Text, 30))
            If sLsnCD = "ALL" Then Exit Sub
            
            '## ������ ������ ����
            
            sSchCD = Trim(basModule.SchCD)
            .Col = 1:       sSisuCD = Trim(.Text)
            Call Get_Lsn_Detail_Note(sSisuCD, sGGbn, sTcr_CL)     '< �� ����[������.Ž��]/ �� ��
            .Col = 2:       sTeacher = Trim(.Text)
            .Col = 4:       sGwamok = Trim(.Text)
            .Col = 3:       nWTotSisu = .Value
            '.Col = 6:       nWLsnSisu = .Value
            
            '<< ���� ���ɽü� ã�� >>
            nWLsnSisu = 0
            Call Get_CanSisu_Data(sSchCD, sSisuCD, sLsnCD, nWLsnSisu)
            .Col = 6
                Call basFunction.Set_SprType_Numeric(sprWork, 0, -99999, 99999, "", nWLsnSisu)
                If nWLsnSisu <= 0 Then
                    sprWork.ForeColor = &HFF&
                Else
                    sprWork.ForeColor = &H80000008
                End If
                
            
            If nWLsnSisu <= 0 Then
                
                With sprWork
                '## [1] �ʱ�ȭ ##########################################
                    For nWorkRow = 1 To .MaxRows Step 1
                        .Row = nWorkRow
                        For nWorkCol = 11 To .MaxCols Step 1
                            .Col = nWorkCol
                                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                        Next nWorkCol
                    Next nWorkRow
                    .Row = 1:   .Row2 = .MaxRows
                    .Col = 11:  .Col2 = .MaxCols
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                End With
                
                lblStatus.Caption = "���ð����� �ü��� �����ϴ�."
                
            Else
                Select Case sGGbn
                    Case "10", "20", "30"       '< ��,��,��
                        Call WorkTable_Schdule_Checks_KME(nTcrRow, sSchCD, sGGbn, sTcr_CL, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                        
                        
                    Case "40", "50"             '< ��,��
                        Call WorkTable_Schdule_Checks_Tamgu(nTcrRow, sSchCD, sGGbn, sTcr_CL, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                        
                End Select
            End If
            
            
            
            
        End If
    End With
    
End Sub


'## ���� ���ɽü� ã��
Private Sub Get_CanSisu_Data(ByVal aSchCD As String, ByVal aSisuCD As String, ByVal aLsnCD As String, ByRef aCanSisu As Long)

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim nRet        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, SISUCD, LSNCD,"
    sStr = sStr & "         SUM(NVL(T_SISU,0)) AS T_SISU,"
    sStr = sStr & "         SUM(NVL(S_SISU,0)) AS S_SISU,"
    sStr = sStr & "         (SUM(NVL(T_SISU,0))-SUM(NVL(S_SISU,0))) AS CAN_SISU"
    sStr = sStr & "    FROM (SELECT ACID, SISUCD, LSNCD, SISU AS T_SISU, 0 AS S_SISU"
    sStr = sStr & "            From SDTCR11TB"
    sStr = sStr & "           WHERE ACID   = '" & aSchCD & "'"
    sStr = sStr & "             AND SISUCD = " & aSisuCD
    sStr = sStr & "          Union All"
    sStr = sStr & "          SELECT ACID, SISUCD, LSNCD, 0 AS T_SISU, SUM(SISU) AS S_SISU"
    sStr = sStr & "            From SDTRX50TB"
    sStr = sStr & "           WHERE ACID   = '" & aSchCD & "'"
    sStr = sStr & "             AND SISUCD = " & aSisuCD
    sStr = sStr & "           GROUP BY ACID, SISUCD, LSNCD"
    sStr = sStr & "          )"
    sStr = sStr & "   WHERE ACID   = '" & aSchCD & "'"
    sStr = sStr & "     AND SISUCD = " & aSisuCD
    sStr = sStr & "     AND LSNCD  = '" & aLsnCD & "'"
    sStr = sStr & "   GROUP BY ACID, SISUCD, LSNCD"
    
    
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
        aCanSisu = 0
        
        If .RecordCount > 0 Then
            .MoveFirst
            If .RecordCount = 1 Then
                If IsNumeric(.Fields("CAN_SISU")) = True Then nRet = CLng(.Fields("CAN_SISU"))
                
                .MoveNext
            End If
        End If
    End With
    
    aCanSisu = nRet
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    aCanSisu = nRet
    
    On Error GoTo 0
End Sub



'## ���� ������ ��������
Private Sub Get_Lsn_Detail_Note(ByVal aSisuCD As String, ByRef aGGbn As String, ByRef aTcr_CL As String)

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT TCRGBN, TCR_CL"
    sStr = sStr & "    From SDTCR01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND SISUCD = " & aSisuCD
    
    
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
        aGGbn = ""
        aTcr_CL = ""
        
        If .RecordCount > 0 Then
            .MoveFirst
            If .RecordCount = 1 Then
                If IsNull(.Fields("TCRGBN")) = False Then aGGbn = Trim(.Fields("TCRGBN"))
                If IsNull(.Fields("TCR_CL")) = False Then aTcr_CL = Trim(.Fields("TCR_CL"))
                .MoveNext
            End If
        End If
    End With
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    aGGbn = ""
    aTcr_CL = ""
    
    On Error GoTo 0
End Sub



'## ��.��.�� ������ ��� #############################################################################################################
'## �Ʒ��� �۾�����
Private Sub WorkTable_Schdule_Checks_KME(ByVal aTcrRow As Long, _
                                         ByVal aSchCD As String, _
                                         ByVal aGbn As String, _
                                         ByVal aSelColor As String, _
                                         ByVal aTeacher As String, _
                                         ByVal aGwamok As String, _
                                         ByVal aLsnCD As String, _
                                         ByVal aWTotSisu As Long, _
                                         ByVal aWLsnSisu As Long)

    
    Dim nWorkRow        As Long
    Dim nWorkCol        As Long
    Dim sTmp            As String
    
    Dim bChk            As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sLesson     As String
    Dim sWeeks      As String
    
    On Error GoTo ErrStmt
    
    
    bChk = False
    lblStatus.Caption = ""
    
    
    
    With sprWork
        
        nWorkRow = aTcrRow
        
        
        '## [1] �ʱ�ȭ ##########################################
        .Row = nWorkRow
        For nWorkCol = 11 To .MaxCols Step 1
            .Col = nWorkCol
                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
        Next nWorkCol
        
        .Row = nWorkRow:    .Row2 = .Row
        .Col = 11:      .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        '## [2] �۾����� ########################################
        
                
        '> 1. ��ü ���� ���ɻ��� ---------------------------------------------------------------------------------------------------------------
        .Row = nWorkRow
        For nWorkCol = 11 To .MaxCols Step 1
            .Col = nWorkCol
                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "1")
        Next nWorkCol
        
        .Row2 = .Row
        .Col = 11:      .Col2 = .MaxCols
        .BlockMode = True
            If aSelColor = "" Then
                .BackColor = basModule.WhiteColor
            Else
                .BackColor = CLng(aSelColor)
            End If
            
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
                
                
        '> 2. ���úҴ��� ���� �˻� << ���Ž �κ� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
        sStr = sStr & "   WHERE A.ACID  = B.ACID"
        sStr = sStr & "     AND A.TRXCD = B.TRXCD"
        sStr = sStr & "     AND A.ACID  = '" & aSchCD & "'"
        sStr = sStr & "     AND A.TRXCD LIKE (SELECT LSNTYPE||'%'"
        sStr = sStr & "                         FROM SDLSN01TB"
        sStr = sStr & "                        WHERE ACID  = '" & aSchCD & "'"
        sStr = sStr & "                          AND LSNCD = '" & aLsnCD & "'"
        sStr = sStr & "                       ) "
        
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        If DBRec.RecordCount > 0 Then
        
            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                    End Select
                    
                End If
                    
                DBRec.MoveNext
            Next nRec
            
        End If
        
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        '> 3. ���úҴ��� ���� �˻� << �̹� ������ ���� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    FROM SDTRX50TB"
        sStr = sStr & "   WHERE ACID  = '" & aSchCD & "'"
        sStr = sStr & "     AND LSNCD = '" & aLsnCD & "'"
        
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        If DBRec.RecordCount > 0 Then
        
            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                    End Select
                    
                End If
                    
                DBRec.MoveNext
            Next nRec
            
        End If
                
        
        
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        '> 4. ���úҴ��� ���� �˻� << ���� �����ϰ�� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    From SDTRX50TB"
        sStr = sStr & "   WHERE (ACID, LSNCD, SISUCD)"
        sStr = sStr & "      IN (SELECT A.ACID, B.LSNCD, A.SISUCD"
        sStr = sStr & "            FROM SDTCR01TB A, SDTCR11TB B"
        sStr = sStr & "           Where A.ACID = B.ACID"
        sStr = sStr & "             AND A.SISUCD = B.SISUCD"
        sStr = sStr & "             AND A.ACID   = '" & aSchCD & "'"
        sStr = sStr & "             AND A.TCRNM  = '" & aTeacher & "'"
        sStr = sStr & "          ) "
        
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

        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop


        If DBRec.RecordCount > 0 Then

            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1

                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then

                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))

                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")

                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False

                    End Select

                End If

                DBRec.MoveNext
            Next nRec

        End If
        
        
        '## ������� �̻������ ###
        bChk = True
        lblStatus.Caption = "�۾� ���̺� �ִ� ������ �����Ͻʽÿ�."
                
    End With
    
    
    If bChk = False Then
        '> ó�� �����̹Ƿ� ���󺹱�
        With sprWork
            
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
            
            .Row = nWorkRow:    .Row2 = .Row
            .Col = 11:          .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End With
    End If
    
    
    
    Exit Sub
ErrStmt:
    '> 1. ��ü ���� ���ɻ���
    With sprWork
        .Row = nWorkRow
        For nWorkCol = 1 To .MaxCols Step 1
            .Col = nWorkCol
                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
        Next nWorkCol
        
        .Row = nWorkRow:    .Row2 = .Row
        .Col = 11:          .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
                
    MsgBox "�۾� �ð�ǥ ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�۾� �ð�ǥ ó��"
    
End Sub






'## ��.��Ž ������ ��� ###########################################################################################################
'## �Ʒ��� �۾�����
Private Sub WorkTable_Schdule_Checks_Tamgu(ByVal aTcrRow As Long, _
                                           ByVal aSchCD As String, _
                                           ByVal aGbn As String, _
                                           ByVal aSelColor As String, _
                                           ByVal aTeacher As String, _
                                           ByVal aGwamok As String, _
                                           ByVal aLsnCD As String, _
                                           ByVal aWTotSisu As Long, _
                                           ByVal aWLsnSisu As Long)


    Dim nWorkRow        As Long
    Dim nWorkCol        As Long
    Dim sTmp            As String
    
    Dim bChk            As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sLesson     As String
    Dim sWeeks      As String
    
    On Error GoTo ErrStmt
    
    
    bChk = False
    lblStatus.Caption = ""
    
    
    
    With sprWork
        
        nWorkRow = aTcrRow
        
        
        '## [1] �ʱ�ȭ ##########################################
        
        .Row = nWorkRow
        For nWorkCol = 11 To .MaxCols Step 1
            .Col = nWorkCol
                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
        Next nWorkCol
        
        .Row2 = .Row
        .Col = 11:      .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        
        '## [2] �۾����� ########################################
                
        '> 1. ���ð��� ���� �˻� << ���Ž �κ� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
        sStr = sStr & "   WHERE A.ACID  = B.ACID"
        sStr = sStr & "     AND A.TRXCD = B.TRXCD"
        sStr = sStr & "     AND A.ACID  = '" & aSchCD & "'"
        sStr = sStr & "     AND A.TRXCD LIKE (SELECT LSNTYPE||'%'"
        sStr = sStr & "                         FROM SDLSN01TB"
        sStr = sStr & "                        WHERE ACID  = '" & aSchCD & "'"
        sStr = sStr & "                          AND LSNCD = '" & aLsnCD & "'"
        sStr = sStr & "                       ) "
        
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
                
                
        If DBRec.RecordCount > 0 Then
        
            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "1")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    If aSelColor = "" Then
                                        .BackColor = basModule.WhiteColor
                                    Else
                                        .BackColor = CLng(aSelColor)
                                    End If
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                    End Select
                    
                End If
                    
                DBRec.MoveNext
            Next nRec
            
        End If
        
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        
        
        
        '> 2. ���úҴ��� ���� �˻� << �̹� ������ ���� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    FROM SDTRX50TB"
        sStr = sStr & "   WHERE ACID  = '" & aSchCD & "'"
        sStr = sStr & "     AND LSNCD = '" & aLsnCD & "'"
        
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
                
                
        If DBRec.RecordCount > 0 Then
        
            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                    End Select
                    
                End If
                    
                DBRec.MoveNext
            Next nRec
            
        End If
                
        
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        '> 3. ���úҴ��� ���� �˻� << ���� �����ϰ�� >> -------------------------------------------------------------------------------------------
        sStr = ""
        sStr = sStr & "  SELECT LESSON, WEEKS"
        sStr = sStr & "    From SDTRX50TB"
        sStr = sStr & "   WHERE (ACID, LSNCD, SISUCD)"
        sStr = sStr & "      IN (SELECT A.ACID, B.LSNCD, A.SISUCD"
        sStr = sStr & "            FROM SDTCR01TB A, SDTCR11TB B"
        sStr = sStr & "           Where A.ACID = B.ACID"
        sStr = sStr & "             AND A.SISUCD = B.SISUCD"
        sStr = sStr & "             AND A.ACID   = '" & aSchCD & "'"
        sStr = sStr & "             AND A.TCRNM  = '" & aTeacher & "'"
        sStr = sStr & "          ) "
        
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
                
                
        If DBRec.RecordCount > 0 Then
        
            DBRec.MoveFirst
            For nRec = 1 To DBRec.RecordCount Step 1
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    .Row = nWorkRow
                    Select Case sWeeks      '< ����//       .COL�� ���� - 1) ���� ó��������ġ 2) ���� 3) -1 �� ������ 1���ʹϱ� !!
                        Case "2"
                            .Col = 11 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:       .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "3"
                            .Col = 21 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                        Case "4"
                            .Col = 31 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "5"
                            .Col = 41 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "6"
                            .Col = 51 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "7"
                            .Col = 61 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                        Case "1"
                            .Col = 71 + CLng(sLesson) - 1
                                Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                
                                .Row2 = .Row:        .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.WhiteColor
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                            
                    End Select
                    
                End If
                    
                DBRec.MoveNext
            Next nRec
            
        End If
        
        
        '## ������� �̻������ ###
        bChk = True
        lblStatus.Caption = "�۾� ���̺� �ִ� ������ �����Ͻʽÿ�."
        
        
    End With
    
    
    If bChk = False Then
        '> ó�� �����̹Ƿ� ���󺹱�
        With sprWork
            
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
                        
            .Row2 = .Row
            .Col = 11:   .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End With
    End If
    
    
    
    Exit Sub
ErrStmt:
    '> 1. ��ü ���� ���ɻ���
    With sprWork
        
        .Row = nWorkRow
        For nWorkCol = 1 To .MaxCols Step 1
            .Col = nWorkCol
                Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
        Next nWorkCol
        
        .Row2 = .Row
        .Col = 11:      .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
                
    MsgBox "�۾� �ð�ǥ ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�۾� �ð�ǥ ó��"
    
End Sub


















'>> �ð�ǥ ���
Private Sub cmdWorkTableSave_Click()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double

    Dim nRow_Work   As Long
    Dim nCol_Work   As Long
    
    Dim nCountChk_S As Long
    Dim nEXE        As Integer
    Dim nAccExe     As Long
    Dim nTotExe     As Long
    
    ReDim uWorkTimeTable(0) As tWorkTimeTable           '< ����� �ڷ�
    
    On Error GoTo ErrStmt
    
    With sprWork
        nCountChk_S = 0     '< S�� üũ�Ǿ��� ����
        
        For nRow_Work = 1 To .MaxRows Step 1
            For nCol_Work = 11 To .MaxCols Step 1
                .Row = nRow_Work
                .Col = nCol_Work
                
                If StrComp(Trim(.Text), "S", vbTextCompare) = 0 Then
                    
                    .Col = 6
                    If .Value > 0 Then      '<< ���ð��� �ü� ���
                    
                        nCountChk_S = nCountChk_S + 1
                        
                        ReDim Preserve uWorkTimeTable(nCountChk_S) As tWorkTimeTable
                        
                        '## ����� ������ ----------------------------------------------------------------
                        
                        uWorkTimeTable(nCountChk_S).ACID = Trim(basModule.SchCD)            '< �п�
                        .Row = nRow_Work
                            .Col = 7:
                                uWorkTimeTable(nCountChk_S).LSNCD = Trim(Right(.Text, 30))  '< ��
                        .Row = SpreadHeader + 2
                            .Col = nCol_Work
                                uWorkTimeTable(nCountChk_S).LESSON = Trim(.Text)            '< ����
                        .Row = SpreadHeader + 1
                            .Col = nCol_Work
                                uWorkTimeTable(nCountChk_S).WEEK = Trim(.Text)              '< ����
                        
                        .Row = nRow_Work
                            .Col = 1
                                uWorkTimeTable(nCountChk_S).SISUCD = Trim(.Text)            '< �ü��ڵ�
                        uWorkTimeTable(nCountChk_S).SISU = "1"                              '< �ü�
                        .Row = nRow_Work
                            .Col = 5
                                uWorkTimeTable(nCountChk_S).TRX_CL = Trim(.BackColor)       '< ��
                        '---------------------------------------------------------------------------------
                        
                        .SetCellBorder nCol_Work, nRow_Work, nCol_Work, nRow_Work, 16, basModule.GridColor2, CellBorderStyleSolid
                        
                    End If
                End If
            Next nCol_Work
        Next nRow_Work
    End With


    If UBound(uWorkTimeTable) = 0 Then  '< S �� ���õ� ������ �����ϴ�.
        MsgBox "����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ ���"
        Exit Sub
    End If
    
    
    nEXE = 0
    nAccExe = 0
    nTotExe = 0
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    
    basDataBase.DBConn.BeginTrans
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    For nRec = 1 To UBound(uWorkTimeTable) Step 1
    
        nTotExe = nTotExe + 1           '<< ó���� ��
        
    
        '>> ��ϵ� ������ ���� ��ȸ
        sStr = ""
        sStr = sStr & "  SELECT ACID, LSNCD, LESSON, WEEKS "
        sStr = sStr & "    FROM SDTRX50TB "
        sStr = sStr & "   WHERE ACID   = '" & uWorkTimeTable(nRec).ACID & "'"
        sStr = sStr & "     AND LSNCD  = '" & uWorkTimeTable(nRec).LSNCD & "'"
        sStr = sStr & "     AND LESSON =  " & uWorkTimeTable(nRec).LESSON
        sStr = sStr & "     AND WEEKS  =  " & uWorkTimeTable(nRec).WEEK
        
        Set DBRec = New ADODB.Recordset
    
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
        
        
    '/* ����ϱ� */
        If DBRec.RecordCount = 0 Then   '<< insert
                
                sStr = ""
                sStr = sStr & "  INSERT INTO SDTRX50TB (ACID, LSNCD, LESSON, WEEKS, SISUCD, SISU, TRX_CL) "
                sStr = sStr & "  VALUES ("
                sStr = sStr & "          '" & uWorkTimeTable(nRec).ACID & "',"
                sStr = sStr & "          '" & uWorkTimeTable(nRec).LSNCD & "',"
                sStr = sStr & "           " & uWorkTimeTable(nRec).LESSON & " ,"
                sStr = sStr & "           " & uWorkTimeTable(nRec).WEEK & " ,"
                sStr = sStr & "           " & uWorkTimeTable(nRec).SISUCD & " ,"
                sStr = sStr & "           " & uWorkTimeTable(nRec).SISU & " ,"
                sStr = sStr & "           " & uWorkTimeTable(nRec).TRX_CL
                sStr = sStr & "  )"
                
    '/* �����ϱ� */
        Else                            '<< update
                sStr = ""
                sStr = sStr & "  UPDATE SDTRX50TB "
                sStr = sStr & "     SET SISUCD =  " & uWorkTimeTable(nRec).SISUCD & " ,"
                sStr = sStr & "         SISU   =  " & uWorkTimeTable(nRec).SISU & " ,"
                sStr = sStr & "         TRX_CL =  " & uWorkTimeTable(nRec).TRX_CL
                
                sStr = sStr & "   WHERE ACID   = '" & uWorkTimeTable(nRec).ACID & "'"
                sStr = sStr & "     AND LSNCD  = '" & uWorkTimeTable(nRec).LSNCD & "'"
                sStr = sStr & "     AND LESSON =  " & uWorkTimeTable(nRec).LESSON
                sStr = sStr & "     AND WEEKS  =  " & uWorkTimeTable(nRec).WEEK
        End If
        Set DBRec = Nothing
        
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
    
        DBCmd.Execute nEXE, , -1
                
                
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nEXE = 1 Then
            nAccExe = nAccExe + 1
        End If
        
    Next nRec
    
    If nTotExe = nAccExe Then
        basDataBase.DBConn.CommitTrans
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    
    '## ���� �ٽ� ��ȸ <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    cmdFind_TeacherData.Tag = "SAVE"
        Call cmdFind_TeacherData_Click
    cmdFind_TeacherData.Tag = ""
    
    cmdShowTimeTable.Tag = "SAVE"
        Call cmdShowTimeTable_Click
    cmdShowTimeTable.Tag = ""
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    If nTotExe = nAccExe Then
        MsgBox "�ð�ǥ ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ð�ǥ ���"
    Else
        MsgBox "�ð�ǥ ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ���"
    End If
    
    Exit Sub
ErrStmt:

    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    MsgBox "�ð�ǥ ��Ͻ� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�ð�ǥ ���"
    
    On Error GoTo 0
    
End Sub




'## ��ϵ� �ð�ǥ ���� ����
Private Sub cmdDelTimeTable_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    Dim nEXE        As Integer
    
    Dim sTmp        As String

    Dim sAcID       As String
    Dim sLsnCD      As String
    Dim sLesson     As String
    Dim sWeeks      As String
    
    Dim sTcrNM      As String
    Dim sSubjNM     As String
    
    
    On Error GoTo ErrStmt
    
    With sprTimeTable
        If .ActiveCol < 1 Then
            MsgBox "������ ������ �����Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ���� ����"
            Exit Sub
        End If
        
        If .ActiveRow < 1 Then
            MsgBox "������ ������ �����Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ���� ����"
            Exit Sub
        End If
        
        '## ��ü���� ��� ��ȸ
        .Row = .ActiveRow
        .Col = SpreadHeader:        sTcrNM = Trim(.Text)
        .Col = .ActiveCol:          sSubjNM = Replace(Trim(.Text), vbCrLf, " ~ ", 1, -1, vbTextCompare)
        
        If MsgBox("���硼 " & sTcrNM & " ��" & vbCrLf & _
                  "���� " & sSubjNM & " �������� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "�ð�ǥ ���û���") = vbNo Then
            Exit Sub
        End If
        
        '## ������ ������
        sAcID = Trim(basModule.SchCD)
        .Col = .ActiveCol
            .Row = SpreadHeader + 1
                sWeeks = Trim(.Text)
            .Row = SpreadHeader + 2
                sLesson = Trim(.Text)
        
        basDataBase.DBConn.BeginTrans
        
            
        sStr = ""
        sStr = sStr & "  DELETE"
        sStr = sStr & "    FROM SDTRX50TB"
        sStr = sStr & "   WHERE (ACID, LSNCD, LESSON, WEEKS)"
        sStr = sStr & "       = (SELECT ACID, LSNCD, LESSON, WEEKS"
        sStr = sStr & "            FROM SDTRX50TB"
        sStr = sStr & "           WHERE ACID = '" & sAcID & "'"
        sStr = sStr & "             AND (ACID, LSNCD, SISUCD)"
        sStr = sStr & "              IN (SELECT A.ACID, B.LSNCD, A.SISUCD"
        sStr = sStr & "                    FROM SDTCR01TB A, SDTCR11TB B"
        sStr = sStr & "                   WHERE A.ACID   = B.ACID"
        sStr = sStr & "                     AND A.SISUCD = B.SISUCD"
        sStr = sStr & "                     AND A.ACID   = '" & sAcID & "'"
        sStr = sStr & "                     AND A.TCRNM  = '" & sTcrNM & "'"
        sStr = sStr & "                   GROUP BY A.ACID, A.SISUCD, B.LSNCD"
        sStr = sStr & "                  )"
        sStr = sStr & "             AND LESSON = " & sLesson
        sStr = sStr & "             AND WEEKS  = " & sWeeks
        sStr = sStr & "          )"
        
        Set DBCmd = New ADODB.Command
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
                
    '    '>> ACID
    '    sTmp = Trim(basModule.SchCD)
    '    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nEXE, , -1
                
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nEXE = 1 Then
            basDataBase.DBConn.CommitTrans
            
            
            '## ���� �ٽ� ��ȸ <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            cmdFind_TeacherData.Tag = "SAVE"
                Call cmdFind_TeacherData_Click
            cmdFind_TeacherData.Tag = ""
            
            cmdShowTimeTable.Tag = "SAVE"
                Call cmdShowTimeTable_Click
            cmdShowTimeTable.Tag = ""
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            MsgBox "�����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ð�ǥ ���û���"
            
        Else
            basDataBase.DBConn.RollbackTrans
            MsgBox "���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ���û���"
        End If
    End With
    
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    On Error Resume Next
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    
    MsgBox "���� ������ ������ �߻��Ͽ����ϴ�." & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�ð�ǥ ���û���"
    
    On Error GoTo 0
End Sub


