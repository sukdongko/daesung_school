VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMR055 
   Caption         =   "시간표 만들기 >> 전체시간표 구성 - 강사별"
   ClientHeight    =   10815
   ClientLeft      =   75
   ClientTop       =   1980
   ClientWidth     =   19095
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   19095
   WindowState     =   2  '최대화
   Begin VB.Frame Frame5 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame5"
      Height          =   6195
      Left            =   60
      TabIndex        =   9
      Top             =   6090
      Width           =   18945
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame4"
         Height          =   6135
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   18885
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "시간표 크게보기"
            Height          =   210
            Index           =   0
            Left            =   1740
            TabIndex        =   17
            Top             =   330
            Width           =   1905
         End
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "시간표 작게보기"
            Height          =   210
            Index           =   1
            Left            =   1740
            TabIndex        =   16
            Top             =   60
            Width           =   1905
         End
         Begin VB.CommandButton cmdDelTimeTable 
            Caption         =   "시간표 내역 삭제"
            Height          =   500
            Left            =   6060
            TabIndex        =   12
            Top             =   30
            Width           =   2595
         End
         Begin VB.CommandButton cmdShowTimeTable 
            Caption         =   "전체시간표 조회"
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
            BackStyle       =   0  '투명
            Caption         =   "전체 시간표"
            BeginProperty Font 
               Name            =   "굴림"
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
      BorderStyle     =   0  '없음
      Caption         =   "Frame3"
      Height          =   5985
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   19005
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   5925
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   18945
         Begin VB.CommandButton cmdFind_TeacherData 
            Caption         =   "강사별 내용 조회"
            Height          =   495
            Left            =   3780
            TabIndex        =   5
            Top             =   90
            Width           =   1845
         End
         Begin VB.CommandButton cmdWorkTableSave 
            Caption         =   "전체 시간표에 반영하기 (시간표 저장)"
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
            BackStyle       =   0  '투명
            Caption         =   "등록 강사의 반별 선택가능 시수내역을 클릭 후  S 를 넣으시면 강제 입력됩니다."
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   2
            Left            =   10530
            TabIndex        =   18
            Top             =   390
            Width           =   7035
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "작업 시간표 테이블"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "lblStatus"
            BeginProperty Font 
               Name            =   "굴림"
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
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM055
'   모 듈  목 적 : 전체시간표 구성 - 강사별
'
'   작   성   일 : 2007/11/28
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
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




'## sprWork 의 spread 헤더 만들기
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
        .Col = 1:       .Text = "시수코드":         .ColWidth(.Col) = 8:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 2:       .Text = "강사":             .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 3:       .Text = "총 시수":          .ColWidth(.Col) = 4:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 4:       .Text = "과목":             .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 5:       .Text = "색":               .ColWidth(.Col) = 5:        .AddCellSpan .Col, .Row, 1, 3
        .Col = 6:       .Text = "가능시수":         .ColWidth(.Col) = 5:        .AddCellSpan .Col, .Row, 1, 3
        
        '<< 반 내역 combo box >>
        .Col = 7:       .Text = "반":               .ColWidth(.Col) = 11:       .AddCellSpan .Col, .Row, 1, 3
        .Col = 8:       .Text = "선택":             .ColWidth(.Col) = 4:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 9:       .Text = " ":                .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        .Col = 10:      .Text = " ":                .ColWidth(.Col) = 6:        .AddCellSpan .Col, .Row, 1, 3:      .ColHidden = True
        
        
        '<< 요일 만들기 >>
        For nCols = 1 To 7 Step 1
            Select Case nCols
                Case 1
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "월"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "2"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 2
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "화"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "3"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 3
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "수"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "4"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 4
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "목"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "5"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 5
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "금"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "6"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 6
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "토"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "7"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 7
                    .Col = nCols * 10 + 1:      .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "일"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
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
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    'XXX
    
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        sprWork.MaxRows = 0
        
        

        '>> 데이터 넣기 --------------------------------------------------------------------
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                
                sprWork.MaxRows = sprWork.MaxRows + 1
                sprWork.Row = sprWork.MaxRows:      sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                
                sprWork.Col = 1:                    sTmp = ""
                    sSisuCD = ""
                    If IsNull(.Fields("SISUCD")) = False Then
                        sTmp = Trim(.Fields("SISUCD")):     sSisuCD = sTmp      '< 시수코드
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
                
                '<< 반내역 조회 >>
                sprWork.Col = sprWork.Col + 1:      sRet = ""
                    
                    '  -- test --
                        sRet = "<<반 선택>>[T]ALL[N]"
                        sRet = sRet & Get_SisuCD_to_Lsn(sSisuCD)
                    
'                        sRet = "인문1" & "[T]" & "00001" & "[N]" & _
'                               "인문2" & "[T]" & "00002" & "[N]" & _
'                               "인문3" & "[T]" & "00003" & "[N]"
                                                  
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
                    
                
                '선택
                sprWork.Col = sprWork.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprWork):       sprWork.Value = 0
                
                '공란
                sprWork.Col = sprWork.Col + 1
                
                '<< 요일내역 >>
                
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
        
        '>> 초기화   -----------------------------------------------------------------------
        sprWork.SetCellBorder 6, 1, 6, sprWork.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        sprWork.SetCellBorder 7, 1, 7, sprWork.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        
        For nWorkRow = 1 To sprWork.MaxRows Step 1
            sprWork.Row = nWorkRow
            For nWorkCol = 10 To sprWork.MaxCols Step 1            '< 요일 내역부터 CLEAR
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
        MsgBox "각 강사별 반을 선택하세요." & vbCrLf & _
               "반 선택시 등록 가능한 시간표 내역을 볼 수 있습니다.", vbInformation + vbOKOnly, "강사시수내역 조회"
               
        cmdFind_TeacherData.Tag = ""
    End If
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    MsgBox "강사별 총시수 내역 조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "강사시수내역 조회"
    
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
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
    
    MsgBox "반 조회시 에러입니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "반 조회"
    
    Get_SisuCD_to_Lsn = sRet
    On Error GoTo 0
End Function




'>> 색 등록
Private Sub sprWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nColor      As Long
    
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
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
            
            '## 취소시엔 CancelColor 로 넘어간다.
        End With
        
        On Error GoTo 0
        On Error GoTo ErrStmt
        
        
        sSchCD = Trim(basModule.SchCD)                                                      '< 학원
        sprWork.Row = Row:      sprWork.Col = 1:        sSisuCD = Trim(sprWork.Text)        '< 시수코드
        
        
        basDataBase.DBConn.BeginTrans

        Set DBCmd = New ADODB.Command
        Set DBParam = New ADODB.Parameter
    
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        
        sStr = ""
        sStr = sStr & "  UPDATE SDTCR01TB"
        sStr = sStr & "     SET TCR_CL =  " & Trim(CStr(nColor))
        sStr = sStr & "   WHERE ACID   = '" & sSchCD & "'"
        sStr = sStr & "     AND SISUCD =  " & sSisuCD
        
        
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
    '    '>> color
    '        nTmp = aColor
    '            Set DBParam = DBCmd.CreateParameter("TRX_CL", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '    '>> 학원
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
            
            MsgBox "색상을 등록하였습니다.", vbInformation + vbOKOnly, "색상 선택하기"
        Else
            basDataBase.DBConn.RollbackTrans
            
            sprWork.Row2 = sprWork.Row:
            sprWork.Col = Col:      sprWork.Col2 = sprWork.Col
            sprWork.BlockMode = True
                sprWork.BackColor = basModule.WhiteColor
                sprWork.BackColorStyle = BackColorStyleUnderGrid
            sprWork.BlockMode = False
            
            MsgBox "색상 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "색상 선택하기"
            
        End If
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
CancelColor:
    MsgBox "선택취소하였습니다.", vbExclamation + vbOKOnly, "색상 선택하기"
    Exit Sub
    
ErrStmt:
    MsgBox "색상 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "색상 선택하기"
    
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
        
        .Col = SpreadHeader:        .Text = "강사":         .ColWidth(.Col) = 6:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 1:    .Text = "담임":         .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 2:    .Text = "총  시수":     .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        .Col = SpreadHeader + 3:    .Text = "선택  시수":   .ColWidth(.Col) = 4:    .AddCellSpan .Col, .Row, 1, 3
        
        
        '<< 요일 만들기 >>
        For nCols = 1 To 7 Step 1
            Select Case nCols
                Case 1
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "월"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "2"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 2
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "화"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "3"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 3
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "수"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "4"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 4
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "목"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "5"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 5
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "금"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "6"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 6
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "토"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
                        For nTmp = 1 To 10 Step 1
                            .Row = SpreadHeader + 1:     .Text = "7"
                            .Row = SpreadHeader + 2:     .Text = Trim(CStr(nTmp))
                            
                            .Col = .Col + 1:    .ColWidth(.Col) = nTtColWidth
                        Next nTmp
                Case 7
                    .Col = (nCols - 1) * 10 + 1:    .ColWidth(.Col) = nTtColWidth
                        .Row = SpreadHeader:         .Text = "일"
                        .AddCellSpan .Col, .Row, 10, 1
                        
                        '## column은 정해진 상태에서 처리
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

' 전체 시간표 조회 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
    
    
    '## 전체내역 모두 조회
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
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
                
''>> 분원
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
                
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    
    If DBRec.RecordCount > 0 Then
    
        DBRec.MoveFirst
        For nRec = 1 To DBRec.RecordCount Step 1
                        
            '> 강사명으로 처리
            nTcrRow = 0
            If IsNull(DBRec.Fields("TCRNM")) = False Then
                sTcrNM = Trim(DBRec.Fields("TCRNM"))
                
                For nRow = 1 To sprTimeTable.MaxRows Step 1     '< 기존 강사명 조회
                    sprTimeTable.Row = nRow
                    sprTimeTable.Col = SpreadHeader
                    If StrComp(Trim(sprTimeTable.Text), sTcrNM, vbTextCompare) = 0 Then
                        nTcrRow = nRow                          '< 강사가 위치한 row
                        
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
                
                If nTcrRow = 0 Then       '>> 강사내역 추가
                    sprTimeTable.MaxRows = sprTimeTable.MaxRows + 1
                    sprTimeTable.Row = sprTimeTable.MaxRows:        sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                    
                    
                    nTcrRow = sprTimeTable.Row                  '< 새로운 row추가된 강사
                    
                    sprTimeTable.Col = SpreadHeader
                        sprTimeTable.Text = Trim(DBRec.Fields("TCRNM"))
                    
                    sprTimeTable.Col = SpreadHeader + 1
                        sprTimeTable.Text = ""                  '< 초기화
                    sprTimeTable.Col = SpreadHeader + 2
                        sprTimeTable.Text = ""                  '< 초기화
                    sprTimeTable.Col = SpreadHeader + 3
                        sprTimeTable.Text = ""                  '< 초기화
                        
                        
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
                
                
            '<< 데이터 넣음 ==================================================================================
                
                ' nTcrRow <- 강사 위치한 row
                
                If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                    sLesson = Trim(DBRec.Fields("LESSON"))
                    sWeeks = Trim(DBRec.Fields("WEEKS"))
                    
                    sTcr_CL = "":       If IsNull(DBRec.Fields("TCR_CL")) = False Then sTcr_CL = Trim(DBRec.Fields("TCR_CL"))                   ' COLOR
                    sDisp_Text = "":    If IsNull(DBRec.Fields("DISP")) = False Then sDisp_Text = Trim(DBRec.Fields("DISP"))                    ' 보여줄 데이터
                                        If IsNull(DBRec.Fields("SUBJNM")) = False Then sDisp_Text = sDisp_Text & vbCrLf & Trim(DBRec.Fields("SUBJNM"))
                    
                    
                    sprTimeTable.Row = nTcrRow              '< 처리할 row
                    
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
            End If      ' 강사명
            
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
        MsgBox "전체시간표 조회하였습니다.", vbInformation + vbOKOnly, "전체 시간표 조회"
        
        cmdShowTimeTable.Tag = ""
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "전체시간표 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체 시간표 조회"
    
End Sub





'## 색 처리
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
        
        
        
        '>> 1로 선택된 부분을 저장가능상태로 바꾸어 줌.
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
                    Call Get_Lsn_Detail_Note(sSisuCD, sGGbn, sTcr_CL)     '< 반 구분[국영수.탐구]/ 반 색
                
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
    Dim sGGbn     As String       ' 과목형태
    Dim sTcr_CL     As String
    Dim sTeacher    As String
    Dim sGwamok     As String
    Dim sLsnCD      As String
    
    Dim nWTotSisu   As Long
    Dim nWLsnSisu   As Long
    
    Dim nWorkRow    As Long
    Dim nWorkCol    As Long
    
    nTcrRow = Row       '< 작업대상 row
    
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
            
            '## 데이터 있으면 진행
            
            sSchCD = Trim(basModule.SchCD)
            .Col = 1:       sSisuCD = Trim(.Text)
            Call Get_Lsn_Detail_Note(sSisuCD, sGGbn, sTcr_CL)     '< 반 구분[국영수.탐구]/ 반 색
            .Col = 2:       sTeacher = Trim(.Text)
            .Col = 4:       sGwamok = Trim(.Text)
            .Col = 3:       nWTotSisu = .Value
            '.Col = 6:       nWLsnSisu = .Value
            
            '<< 강의 가능시수 찾기 >>
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
                '## [1] 초기화 ##########################################
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
                
                lblStatus.Caption = "선택가능한 시수가 없습니다."
                
            Else
                Select Case sGGbn
                    Case "10", "20", "30"       '< 언,수,외
                        Call WorkTable_Schdule_Checks_KME(nTcrRow, sSchCD, sGGbn, sTcr_CL, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                        
                        
                    Case "40", "50"             '< 사,과
                        Call WorkTable_Schdule_Checks_Tamgu(nTcrRow, sSchCD, sGGbn, sTcr_CL, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                        
                End Select
            End If
            
            
            
            
        End If
    End With
    
End Sub


'## 강의 가능시수 찾기
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
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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



'## 반의 상세정보 가져오기
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
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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



'## 언.수.외 선택인 경우 #############################################################################################################
'## 아래의 작업진행
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
        
        
        '## [1] 초기화 ##########################################
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
        
        
        '## [2] 작업진행 ########################################
        
                
        '> 1. 전체 선택 가능상태 ---------------------------------------------------------------------------------------------------------------
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
                
                
        '> 2. 선택불능인 내용 검색 << 사과탐 부분 >> -------------------------------------------------------------------------------------------
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
                
    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        '> 3. 선택불능인 내용 검색 << 이미 선택한 내용 >> -------------------------------------------------------------------------------------------
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
        
    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        '> 4. 선택불능인 내용 검색 << 같은 강사일경우 >> -------------------------------------------------------------------------------------------
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

        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni

    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열

        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        
        '## 여기까지 이상없으면 ###
        bChk = True
        lblStatus.Caption = "작업 테이블에 있는 내용을 선택하십시요."
                
    End With
    
    
    If bChk = False Then
        '> 처리 오류이므로 원상복귀
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
    '> 1. 전체 선택 가능상태
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
                
    MsgBox "작업 시간표 처리시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "작업 시간표 처리"
    
End Sub






'## 사.과탐 선택인 경우 ###########################################################################################################
'## 아래의 작업진행
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
        
        
        '## [1] 초기화 ##########################################
        
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
        
        
        
        '## [2] 작업진행 ########################################
                
        '> 1. 선택가능 내용 검색 << 사과탐 부분 >> -------------------------------------------------------------------------------------------
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
                
    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        
        
        
        '> 2. 선택불능인 내용 검색 << 이미 선택한 내용 >> -------------------------------------------------------------------------------------------
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
        
    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        '> 3. 선택불능인 내용 검색 << 같은 강사일경우 >> -------------------------------------------------------------------------------------------
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
        
    '    '>> 분원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '    '>> 계열
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                    Select Case sWeeks      '< 요일//       .COL의 내용 - 1) 요일 처음시작위치 2) 교시 3) -1 은 시작이 1부터니깐 !!
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
        
        
        '## 여기까지 이상없으면 ###
        bChk = True
        lblStatus.Caption = "작업 테이블에 있는 내용을 선택하십시요."
        
        
    End With
    
    
    If bChk = False Then
        '> 처리 오류이므로 원상복귀
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
    '> 1. 전체 선택 가능상태
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
                
    MsgBox "작업 시간표 처리시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "작업 시간표 처리"
    
End Sub


















'>> 시간표 등록
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
    
    ReDim uWorkTimeTable(0) As tWorkTimeTable           '< 등록할 자료
    
    On Error GoTo ErrStmt
    
    With sprWork
        nCountChk_S = 0     '< S로 체크되어진 갯수
        
        For nRow_Work = 1 To .MaxRows Step 1
            For nCol_Work = 11 To .MaxCols Step 1
                .Row = nRow_Work
                .Col = nCol_Work
                
                If StrComp(Trim(.Text), "S", vbTextCompare) = 0 Then
                    
                    .Col = 6
                    If .Value > 0 Then      '<< 선택가능 시수 계산
                    
                        nCountChk_S = nCountChk_S + 1
                        
                        ReDim Preserve uWorkTimeTable(nCountChk_S) As tWorkTimeTable
                        
                        '## 등록할 데이터 ----------------------------------------------------------------
                        
                        uWorkTimeTable(nCountChk_S).ACID = Trim(basModule.SchCD)            '< 학원
                        .Row = nRow_Work
                            .Col = 7:
                                uWorkTimeTable(nCountChk_S).LSNCD = Trim(Right(.Text, 30))  '< 반
                        .Row = SpreadHeader + 2
                            .Col = nCol_Work
                                uWorkTimeTable(nCountChk_S).LESSON = Trim(.Text)            '< 교시
                        .Row = SpreadHeader + 1
                            .Col = nCol_Work
                                uWorkTimeTable(nCountChk_S).WEEK = Trim(.Text)              '< 요일
                        
                        .Row = nRow_Work
                            .Col = 1
                                uWorkTimeTable(nCountChk_S).SISUCD = Trim(.Text)            '< 시수코드
                        uWorkTimeTable(nCountChk_S).SISU = "1"                              '< 시수
                        .Row = nRow_Work
                            .Col = 5
                                uWorkTimeTable(nCountChk_S).TRX_CL = Trim(.BackColor)       '< 색
                        '---------------------------------------------------------------------------------
                        
                        .SetCellBorder nCol_Work, nRow_Work, nCol_Work, nRow_Work, 16, basModule.GridColor2, CellBorderStyleSolid
                        
                    End If
                End If
            Next nCol_Work
        Next nRow_Work
    End With


    If UBound(uWorkTimeTable) = 0 Then  '< S 로 선택된 내용이 없습니다.
        MsgBox "등록할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 등록"
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
    
        nTotExe = nTotExe + 1           '<< 처리한 수
        
    
        '>> 등록된 데이터 여부 조회
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
    
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
    '/* 등록하기 */
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
                
    '/* 갱신하기 */
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
        
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
    '    '>> 분원
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
    
    
    '## 전부 다시 조회 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
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
        MsgBox "시간표 등록하였습니다.", vbInformation + vbOKOnly, "시간표 등록"
    Else
        MsgBox "시간표 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 등록"
    End If
    
    Exit Sub
ErrStmt:

    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    MsgBox "시간표 등록시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 등록"
    
    On Error GoTo 0
    
End Sub




'## 등록된 시간표 내역 삭제
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
            MsgBox "삭제할 내용을 선택하여 주십시요.", vbExclamation + vbOKOnly, "시간표 내역 삭제"
            Exit Sub
        End If
        
        If .ActiveRow < 1 Then
            MsgBox "삭제할 내용을 선택하여 주십시요.", vbExclamation + vbOKOnly, "시간표 내역 삭제"
            Exit Sub
        End If
        
        '## 전체내역 모두 조회
        .Row = .ActiveRow
        .Col = SpreadHeader:        sTcrNM = Trim(.Text)
        .Col = .ActiveCol:          sSubjNM = Replace(Trim(.Text), vbCrLf, " ~ ", 1, -1, vbTextCompare)
        
        If MsgBox("강사【 " & sTcrNM & " 】" & vbCrLf & _
                  "과목【 " & sSubjNM & " 】내용을 삭제하시겠습니까?", vbQuestion + vbYesNo, "시간표 선택삭제") = vbNo Then
            Exit Sub
        End If
        
        '## 삭제할 데이터
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
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
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
            
            
            '## 전부 다시 조회 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            cmdFind_TeacherData.Tag = "SAVE"
                Call cmdFind_TeacherData_Click
            cmdFind_TeacherData.Tag = ""
            
            cmdShowTimeTable.Tag = "SAVE"
                Call cmdShowTimeTable_Click
            cmdShowTimeTable.Tag = ""
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            
            MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간표 선택삭제"
            
        Else
            basDataBase.DBConn.RollbackTrans
            MsgBox "삭제 오류가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 선택삭제"
        End If
    End With
    
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    On Error Resume Next
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    
    MsgBox "선택 삭제시 에러가 발생하였습니다." & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 선택삭제"
    
    On Error GoTo 0
End Sub


