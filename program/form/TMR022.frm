VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR022 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "시간표 만들기 >> 반 구성하기 >> 학생신청과목 펼친내역 보기"
   ClientHeight    =   9225
   ClientLeft      =   720
   ClientTop       =   1785
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15075
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   855
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   795
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   14985
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   8460
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   90
            Width           =   1215
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   6570
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   75
            Width           =   1035
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회"
            Height          =   465
            Left            =   210
            TabIndex        =   0
            Top             =   180
            Width           =   1725
         End
         Begin EditLib.fpMask fpExmID1 
            Height          =   300
            Left            =   3000
            TabIndex        =   1
            Top             =   75
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
            Mask            =   "99999"
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
         Begin EditLib.fpMask fpExmID2 
            Height          =   300
            Left            =   4200
            TabIndex        =   2
            Top             =   75
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
            _ExtentY        =   529
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
            Mask            =   "99999"
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
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "유/무시험"
            Height          =   210
            Left            =   5490
            TabIndex        =   11
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "수강가능 처리내역 보기를 하시면 반영된 내역까지 조회가능합니다."
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
            Left            =   2160
            TabIndex        =   10
            Top             =   540
            Width           =   7185
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   7350
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "수험번호               부터               까지"
            Height          =   210
            Left            =   2160
            TabIndex        =   8
            Top             =   150
            Width           =   3675
         End
      End
   End
   Begin FPSpread.vaSpread sprData 
      Height          =   5475
      Left            =   30
      TabIndex        =   5
      Top             =   930
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   9657
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
      MaxCols         =   30
      SpreadDesigner  =   "TMR022.frx":0000
   End
   Begin FPSpread.vaSpread sprLsn 
      Height          =   2775
      Left            =   30
      TabIndex        =   12
      Top             =   6420
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
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
      MaxCols         =   20
      SpreadDesigner  =   "TMR022.frx":20D8
   End
End
Attribute VB_Name = "TMR022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM022
'   모 듈  목 적 : 작업된 학생내역 모두 보여주기
'
'   작   성   일 : 2007/11/13
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit


Private Sub Form_Load()
    
    Me.Move 50, 900, 15195, 9630
    
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
                
        With cboExmType
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "무시험" & Space(30) & "0"
            .AddItem "유시험" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        fpExmID1.Text = ""
        fpExmID2.Text = ""
        
    Me.Tag = ""
    
End Sub

Private Sub cmdFind_Click()
    
    Call Find_STD_Data              '< 학생조회
    Call Find_Lsn_To_STD_TOT        '< 반별별 합계내역
    Call Find_Gwamok_to_STD_TOT     '< 과목별 합계내역
    
    MsgBox "조회하였습니다.", vbInformation + vbOKOnly, "조회"

End Sub


'## 학생의 신청내용 조회
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
    sStr = sStr & "                 EXMTYPE, DECODE(EXMTYPE,'0','무','유') AS EXMTYPE_NM,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                     '01'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                     '02'"
    sStr = sStr & "                 END END GAEYUL_CD,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                     '사탐'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                     '과탐'"
    sStr = sStr & "                 END END GAEYUL,"
    
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN"
    sStr = sStr & "                     '국사'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN"
    sStr = sStr & "                     '물1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL1,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN"
    sStr = sStr & "                     '윤리'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN"
    sStr = sStr & "                     '화1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL2,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN"
    sStr = sStr & "                     '경제'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN"
    sStr = sStr & "                     '생1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL3,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN"
    sStr = sStr & "                     '한근'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN"
    sStr = sStr & "                     '지1'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL4,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN"
    sStr = sStr & "                     '세계사'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN"
    sStr = sStr & "                     '물2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL5,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN"
    sStr = sStr & "                     '경지'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN"
    sStr = sStr & "                     '화2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL6,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN"
    sStr = sStr & "                     '한지'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN"
    sStr = sStr & "                     '생2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL7,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN"
    sStr = sStr & "                     '정치'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN"
    sStr = sStr & "                     '지2'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END SEL8,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN"
    sStr = sStr & "                     '사문'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL9,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN"
    sStr = sStr & "                     '법사'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL10,"
    sStr = sStr & "                 CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN"
    sStr = sStr & "                     '세지'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL11,"
    
    sStr = sStr & "              /* 제2외국어 & 수리 */"
    sStr = sStr & "                      CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독일어'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에스파냐'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중국어'"
    sStr = sStr & "                 ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적분'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산수학'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률통계'"
    sStr = sStr & "                 ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '수리나형'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "              /* 논술 */"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                     '언어'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N1,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                     '수리'"
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N2,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                     '외국어'"                                           '< 변경
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N3,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                     ''"                                                 '< 변경
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
    
    If Trim(fpExmID1.UnFmtText) <> "" And Trim(fpExmID2.UnFmtText) <> "" Then
        sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID1.UnFmtText) & "'"
        sStr = sStr & "                       AND '" & Trim(fpExmID2.UnFmtText) & "'"
    ElseIf Trim(fpExmID1.UnFmtText) <> "" And Trim(fpExmID2.UnFmtText) = "" Then
        sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID1.UnFmtText) & "'"
        sStr = sStr & "                       AND '99999'"
    ElseIf Trim(fpExmID1.UnFmtText) = "" And Trim(fpExmID2.UnFmtText) <> "" Then
        sStr = sStr & "         AND EXMID BETWEEN '00000'"
        sStr = sStr & "                       AND '" & Trim(fpExmID2.UnFmtText) & "'"
    Else
        'no action
    End If
    
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
    sStr = sStr & "   ORDER BY EXMID, GAEYUL_CD, SEL_CLASS, STDNM"
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> EXMID1
'        sTmp = Left(fpExmID1.UnFmtText, 5)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> EXMID2
'        sTmp = Left(fpExmID2.UnFmtText, 5)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprData.MaxRows = sprData.MaxRows + 1
                sprData.Row = sprData.MaxRows

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

                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("GAEYUL_CD")) = False Then sTmp = Trim(.Fields("GAEYUL_CD"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)

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

                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                

                For ni = 1 To 4 Step 1
                    sFieldNM = ""

                    sFieldNM = "SEL_N" & Trim(CStr(ni))
                    sprData.Col = sprData.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni

                
                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS")) = False Then sTmp = Trim(.Fields("SEL_CLASS"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS_NM")) = False Then sTmp = Trim(.Fields("SEL_CLASS_NM"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ":     If IsNull(.Fields("CL_CLOSE")) = False Then sTmp = Trim(.Fields("CL_CLOSE"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)


                sprData.SetCellBorder sprData.Col, sprData.Row, sprData.Col, sprData.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                

                For ni = 1 To 4 Step 1
                    sFieldNM = ""

                    sFieldNM = "GWA_BAN" & Trim(CStr(ni))
                    sprData.Col = sprData.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni
            
                .MoveNext       '<< 다음항목
                
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
    MsgBox "학생 신청내역 조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
    
End Sub



'## 반별 과목신청내역 합계인원
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
    sStr = sStr & "         DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                       '02','자연 J') AS KAEYOL_NM"
    
    sStr = sStr & "    FROM (SELECT LSNCD,"
    sStr = sStr & "                 GET_LSNNM(ACID, LSNCD) AS LSNNM,"
    
    sStr = sStr & "                 COUNT(CL_CLOSE) AS INWON_STAT,                      /* 작업완료 된 학생 */"
    
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
    
    sStr = sStr & "                 COUNT(SEL_X2) AS SEL_X2,"

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
    
    sStr = sStr & "                  FROM (SELECT ACID, "
    sStr = sStr & "                               SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* 사탐, 과탐 분리 */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* 사탐-국사 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* 사탐-윤리 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생물1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* 사탐-한국근현대 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* 사탐-경제지리 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생물2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* 사탐-정치 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* 사탐-사회문화 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* 사탐-법과사회 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* 사탐-세계지리 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    
    sStr = sStr & "                           /* 제2외국어 & 수리 */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독일어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에스파냐'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중국어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적분'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산수학'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률통계'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '수리나형'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                        ''"
    sStr = sStr & "                               END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "                           /* 논술 */"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                                   '언어'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N1,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                                       '수리'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N2,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                                       '외국어'"                             '< 변경
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N3,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                                       ''"                                   '< 변경
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
    sStr = sStr & "              GROUP BY ACID, LSNCD"
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
    


    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 계열
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 계열
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 반 형태
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
            
                nRet = nRet + 1
                
                sprLsn.MaxRows = sprLsn.MaxRows + 1
                sprLsn.Row = sprLsn.MaxRows
                
                sprLsn.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprLsn.Col = sprLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprLsn.Col = sprLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("KAEYOL_NM")) = False Then sTmp = Trim(.Fields("KAEYOL_NM"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '<< 인문자연 공통 : 8 과목 >>
                For nCol = 1 To 8 Step 1
                    sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '사탐은 9~11
                        For nCol = 9 To 11 Step 1
                            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                            
                        Next nCol
                        
                    Case "02"
                        '과탐은 COLUMN만 이동
                        For nCol = 9 To 11 Step 1
                            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                            If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                        Next nCol
                End Select
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 제2선택
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_X2")) = False Then
                        nTmp = CDbl(.Fields("SEL_X2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 언
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N1")) = False Then
                        nTmp = CDbl(.Fields("SEL_N1"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 수
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N2")) = False Then
                        nTmp = CDbl(.Fields("SEL_N2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 사
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N3")) = False Then
                        nTmp = CDbl(.Fields("SEL_N3"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 탐
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N4")) = False Then
                        nTmp = CDbl(.Fields("SEL_N4"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '>> 총인원
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                
                .MoveNext       '<< 다음항목
                
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
    MsgBox "반별 수강신청내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "반별 수강신청내역 조회"
    
    Find_Lsn_To_STD_TOT = nRet
End Function





'## 전체 과목별 학생수
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
    
    sStr = sStr & "    FROM (SELECT COUNT(CL_CLOSE) AS INWON_STAT,                      /* 작업완료 된 학생 */"
    
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
    
    sStr = sStr & "                 COUNT(SEL_X2) AS SEL_X2,"

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
    
    sStr = sStr & "                        /* 사탐, 과탐 분리 */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* 사탐-국사 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* 사탐-윤리 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생물1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* 사탐-한국근현대 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* 사탐-경제지리 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생물2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* 사탐-정치 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* 사탐-사회문화 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* 사탐-법과사회 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* 사탐-세계지리 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    
    sStr = sStr & "                           /* 제2외국어 & 수리 */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독일어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에스파냐'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중국어'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적분'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산수학'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률통계'"
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '수리나형'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                        ''"
    sStr = sStr & "                               END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "                           /* 논술 */"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'91|') > 0 THEN"
    sStr = sStr & "                                   '언어'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N1,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'92|') > 0 THEN"
    sStr = sStr & "                                       '수리'"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N2,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'93|') > 0 THEN"
    sStr = sStr & "                                       '외국어'"                             '< 변경
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N3,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                                       ''"                                   '< 변경
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
    


    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 계열
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 계열
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 반 형태
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            sprLsn.MaxRows = sprLsn.MaxRows + 1
            sprLsn.InsertRows 1, 1
            sprLsn.Row = 1
            
            sprLsn.SetCellBorder 1, sprLsn.Row, sprLsn.MaxCols, sprLsn.Row, 8, basModule.SectionColor1, CellBorderStyleSolid
                
                
            sprLsn.Col = 1
                sTmp = " "
                    Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
            sprLsn.Col = sprLsn.Col + 1
                sTmp = " "
                    Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
            sprLsn.Col = sprLsn.Col + 1
                sTmp = "합 계 "
                    Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprLsn.ForeColor = basModule.SectionColor1
                
            'sprLsn.ForeColor = &H0
            sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '<< 인문자연 공통 : 8 과목 >>
                For nCol = 1 To 11 Step 1
                    sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 제2선택
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_X2")) = False Then
                        nTmp = CDbl(.Fields("SEL_X2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 언
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N1")) = False Then
                        nTmp = CDbl(.Fields("SEL_N1"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 수
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N2")) = False Then
                        nTmp = CDbl(.Fields("SEL_N2"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 사
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N3")) = False Then
                        nTmp = CDbl(.Fields("SEL_N3"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                '> 탐
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N4")) = False Then
                        nTmp = CDbl(.Fields("SEL_N4"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                sprLsn.SetCellBorder sprLsn.Col, sprLsn.Row, sprLsn.Col, sprLsn.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '>> 총인원
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                
                
            
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
    
    Find_Gwamok_to_STD_TOT = nRet
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "각 과목별 수강신청내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "과목별 수강신청내역 조회"
    
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

Private Sub sprLsn_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprLsn
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


















