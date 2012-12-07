VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMR050 
   Caption         =   "시간표 만들기 >> 전체시간표 구성 - 반별"
   ClientHeight    =   11295
   ClientLeft      =   30
   ClientTop       =   2430
   ClientWidth     =   19260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11295
   ScaleWidth      =   19260
   WindowState     =   2  '최대화
   Begin VB.Frame Frame8 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame8"
      Height          =   3795
      Left            =   30
      TabIndex        =   35
      Top             =   7500
      Width           =   19125
      Begin VB.Frame Frame9 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame9"
         Height          =   3735
         Left            =   30
         TabIndex        =   36
         Top             =   30
         Width           =   19065
         Begin VB.CommandButton cmdDelTimeTable 
            Caption         =   "시간표 내역 삭제"
            Height          =   450
            Left            =   5940
            TabIndex        =   8
            Top             =   30
            Width           =   2595
         End
         Begin VB.CommandButton cmdShowTimeTable 
            Caption         =   "전체 시간표 조회"
            Height          =   450
            Left            =   2250
            TabIndex        =   7
            Top             =   30
            Width           =   2595
         End
         Begin FPSpread.vaSpread sprTimeTable 
            Height          =   3135
            Left            =   0
            TabIndex        =   9
            Top             =   540
            Width           =   19065
            _Version        =   393216
            _ExtentX        =   33629
            _ExtentY        =   5530
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
            SpreadDesigner  =   "TMR050.frx":0000
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
            Left            =   180
            TabIndex        =   38
            Top             =   150
            Width           =   3075
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "삭제할 내역 선택 후 클릭하세요."
            ForeColor       =   &H001E5A75&
            Height          =   210
            Index           =   1
            Left            =   8610
            TabIndex        =   37
            Top             =   150
            Width           =   2805
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame5"
      Height          =   2865
      Left            =   0
      TabIndex        =   29
      Top             =   4590
      Width           =   19155
      Begin VB.Frame Frame6 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame6"
         Height          =   2805
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   19095
         Begin VB.CommandButton cmdWorkTableSave 
            Caption         =   "전체 시간표에 반영하기 (시간표 저장)"
            Height          =   465
            Left            =   11490
            TabIndex        =   6
            Top             =   2280
            Width           =   3945
         End
         Begin FPSpread.vaSpread sprWork 
            Height          =   1935
            Left            =   0
            TabIndex        =   5
            Top             =   300
            Width           =   19065
            _Version        =   393216
            _ExtentX        =   33629
            _ExtentY        =   3413
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
            SpreadDesigner  =   "TMR050.frx":01D4
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "등록 강사의 반별 선택가능 시수내역을 클릭 후  S 를 넣으시면 강제 입력됩니다."
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   39
            Top             =   2490
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
            TabIndex        =   33
            Top             =   60
            Width           =   3075
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "마우스 왼쪽 두번 클릭시에 기본교실 & 담임교사를 넣을 수 있습니다."
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   32
            Top             =   2280
            Width           =   7035
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
            Left            =   2520
            TabIndex        =   31
            Top             =   60
            Width           =   6405
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame4"
      Height          =   4485
      Left            =   30
      TabIndex        =   25
      Top             =   60
      Width           =   11985
      Begin VB.Frame Frame3 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame3"
         Height          =   4425
         Left            =   30
         TabIndex        =   26
         Top             =   30
         Width           =   5025
         Begin VB.Frame Frame7 
            BackColor       =   &H001E5A75&
            BorderStyle     =   0  '없음
            Caption         =   "Frame7"
            Height          =   15
            Left            =   120
            TabIndex        =   34
            Top             =   810
            Width           =   4785
         End
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "시간표 크게보기"
            Height          =   210
            Index           =   0
            Left            =   2160
            TabIndex        =   2
            Top             =   600
            Width           =   1905
         End
         Begin VB.OptionButton optView 
            BackColor       =   &H00D2EAF5&
            Caption         =   "시간표 작게보기"
            Height          =   210
            Index           =   1
            Left            =   270
            TabIndex        =   1
            Top             =   600
            Width           =   1905
         End
         Begin VB.CommandButton cmdTotSisu 
            Caption         =   "강사/과목별 시수조회"
            Height          =   400
            Left            =   270
            TabIndex        =   0
            Top             =   60
            Width           =   2205
         End
         Begin FPSpread.vaSpread sprTotSisu 
            Height          =   3285
            Left            =   90
            TabIndex        =   3
            Top             =   870
            Width           =   4815
            _Version        =   393216
            _ExtentX        =   8493
            _ExtentY        =   5794
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
            SpreadDesigner  =   "TMR050.frx":03A8
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "색 컬럼을 두번클릭시에 강사별 과목의 색을 지정할 수 있음."
            Height          =   210
            Left            =   90
            TabIndex        =   27
            Top             =   4200
            Width           =   4995
         End
      End
      Begin FPSpread.vaSpread sprLsnSisu 
         Height          =   3825
         Left            =   5100
         TabIndex        =   4
         Top             =   630
         Width           =   6855
         _Version        =   393216
         _ExtentX        =   12091
         _ExtentY        =   6747
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
         MaxCols         =   11
         SpreadDesigner  =   "TMR050.frx":1E54
      End
      Begin VB.Label Label24 
         BackStyle       =   0  '투명
         Caption         =   ">> 강사의 반별 선택가능 시수내역 조회"
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
         Height          =   285
         Left            =   5130
         TabIndex        =   28
         Top             =   390
         Width           =   6435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   5475
      Left            =   12180
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   3555
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   1500
         Width           =   3075
         Begin VB.TextBox txtSelLsnCD 
            Height          =   300
            Left            =   840
            TabIndex        =   18
            Text            =   "txtSelLsnCD"
            Top             =   1440
            Width           =   1905
         End
         Begin VB.TextBox txtSelSisuCD 
            Height          =   300
            Left            =   840
            TabIndex        =   17
            Text            =   "txtSelSisuCD"
            Top             =   450
            Width           =   1905
         End
         Begin VB.TextBox txtSelSchCD 
            Height          =   300
            Left            =   840
            TabIndex        =   16
            Text            =   "txtSelSchCD"
            Top             =   150
            Width           =   1905
         End
         Begin VB.TextBox txtnColor 
            Height          =   300
            Left            =   1530
            TabIndex        =   15
            Text            =   "txtnColor"
            Top             =   1110
            Width           =   1215
         End
         Begin VB.TextBox txtSelGbn 
            Height          =   300
            Left            =   840
            TabIndex        =   14
            Text            =   "txtSelGbn"
            Top             =   750
            Width           =   1905
         End
         Begin EditLib.fpLongInteger fpWorkSisu 
            Height          =   300
            Left            =   840
            TabIndex        =   19
            Top             =   1110
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
         Begin EditLib.fpLongInteger fpsprLsnSisuRow 
            Height          =   300
            Left            =   1920
            TabIndex        =   20
            Top             =   1770
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
         Begin VB.Label Label6 
            Caption         =   "sprLsnSisuRow"
            Height          =   210
            Left            =   840
            TabIndex        =   22
            Top             =   1830
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "작업시수"
            Height          =   210
            Left            =   60
            TabIndex        =   21
            Top             =   1110
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdFindLsnSisu 
         Caption         =   "반의 선택시수 조회"
         Height          =   400
         Left            =   120
         TabIndex        =   12
         Top             =   1050
         Width           =   2145
      End
      Begin VB.CommandButton cmdFindWork 
         Caption         =   "작업sheet 조회"
         Height          =   400
         Left            =   120
         TabIndex        =   11
         Top             =   3750
         Width           =   2355
      End
      Begin EditLib.fpLongInteger fpsprTotSisu_Row 
         Height          =   300
         Left            =   1260
         TabIndex        =   23
         Top             =   330
         Width           =   675
         _Version        =   196608
         _ExtentX        =   1191
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
         BackColor       =   12632256
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
      Begin VB.Label Label8 
         Caption         =   "sprTotSisu"
         Height          =   210
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   975
      End
   End
End
Attribute VB_Name = "TMR050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM050
'   모 듈  목 적 : 전체시간표 구성
'
'   작   성   일 : 2007/11/16
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
    
    
    With sprTotSisu
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
    End With
    
    With sprLsnSisu
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
    End With
    
    With sprWork
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
        .MaxCols = 0
    End With
    
    With sprTimeTable
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
        .MaxCols = 0
    End With
    
    Me.Tag = "LOAD"
        
        cmdTotSisu.Tag = ""
        
        fpsprTotSisu_Row.Value = 0
        
        lblStatus.Caption = ""
        txtnColor.Text = ""
        
        txtSelSchCD.Text = ""
        txtSelSisuCD.Text = ""
        fpsprLsnSisuRow.Value = 0
        
        optView(0).Value = False
        optView(1).Value = True
        
        If optView(0).Value = True Then
            nTtRowHeight = 25
            nTtColWidth = 6
        ElseIf optView(1).Value = True Then
            nTtRowHeight = 15
            nTtColWidth = 3
        End If
        
    Me.Tag = ""
    
End Sub






'## 반 COUNT
Private Function Find_LsnCount() As Long
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim nRet        As Long
    
    nRet = 0
    On Error GoTo ErrStmt
    
    
    sStr = ""
    sStr = sStr & "  SELECT COUNT(LSNCD) AS LSNCD_CNT"
    sStr = sStr & "    From SDLSN01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        
    
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
        If .RecordCount > 0 Then
            .MoveFirst
                
            If IsNull(.Fields("LSNCD_CNT")) = False Then
                nRet = CLng(.Fields("LSNCD_CNT"))
                
            End If
        End If
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    Find_LsnCount = nRet
    
    Exit Function
ErrStmt:
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
    
    Find_LsnCount = nRet
    
End Function


'=====================================================================================================
' 계열에 속한 모든 반에 대한 내용을 display하게 됨.
'=====================================================================================================
Private Sub Construct_Base_TimeTable(ByVal aLsnCount As Long)
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCols       As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    '/* cols & rows 조정 */
    If optView(0).Value = True Then
        nTtRowHeight = 25
        nTtColWidth = 6
    ElseIf optView(1).Value = True Then
        nTtRowHeight = 20
        nTtColWidth = 2
    End If
    
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, BASE_CLASS, DAMIM"
    sStr = sStr & "    From SDLSN01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "   ORDER BY KAEYOL, LSNNM "
    
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
        If .RecordCount > 0 Then
            .MoveFirst
            
            sprWork.Col = SpreadHeader:         sprWork.ColHidden = False
            sprWork.MaxRows = aLsnCount
            sprWork.RowHeaderCols = 4
            sprWork.MaxCols = 0
            sprWork.ColHeaderRows = 3
            sprWork.Col = SpreadHeader:         sprWork.ColHidden = True
            sprWork.Row = SpreadHeader + 1:     sprWork.RowHidden = True
            
            sprWork.Row = SpreadHeader
                sprWork.Col = SpreadHeader:     sprWork.Text = "반코드":    sprWork.AddCellSpan sprWork.Col, sprWork.Row, 1, 3
                sprWork.Col = SpreadHeader + 1: sprWork.Text = "반":        sprWork.AddCellSpan sprWork.Col, sprWork.Row, 1, 3
                sprWork.Col = SpreadHeader + 2: sprWork.Text = "기본반":    sprWork.AddCellSpan sprWork.Col, sprWork.Row, 1, 3
                sprWork.Col = SpreadHeader + 3: sprWork.Text = "담임":      sprWork.AddCellSpan sprWork.Col, sprWork.Row, 1, 3
            
            
            sprTimeTable.Col = SpreadHeader:        sprTimeTable.ColHidden = False
            sprTimeTable.MaxRows = aLsnCount
            sprTimeTable.RowHeaderCols = 4
            sprTimeTable.MaxCols = 0
            sprTimeTable.ColHeaderRows = 3
            sprTimeTable.Col = SpreadHeader:        sprTimeTable.ColHidden = True
            sprTimeTable.Row = SpreadHeader + 1:    sprTimeTable.RowHidden = True
            
            sprTimeTable.Row = SpreadHeader
                sprTimeTable.Col = SpreadHeader:     sprTimeTable.Text = "반코드":    sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
                sprTimeTable.Col = SpreadHeader + 1: sprTimeTable.Text = "반":        sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
                sprTimeTable.Col = SpreadHeader + 2: sprTimeTable.Text = "기본반":    sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
                sprTimeTable.Col = SpreadHeader + 3: sprTimeTable.Text = "담임":      sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
            
            
            For nRec = 1 To .RecordCount Step 1
            
                '<< 작업테이블 >>
                    sprWork.Col = SpreadHeader:             sprWork.ColWidth(sprWork.Col) = nTtColWidth
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 1:         sprWork.ColWidth(sprWork.Col) = 6
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 2:         sprWork.ColWidth(sprWork.Col) = 4
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("BASE_CLASS")) = False Then sTmp = Trim(.Fields("BASE_CLASS"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 3:         sprWork.ColWidth(sprWork.Col) = 6
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("DAMIM")) = False Then sTmp = Trim(.Fields("DAMIM"))
                        sprWork.Text = sTmp
                        
                '<< 요일 만들기 >>
                sprWork.MaxCols = 70
                For nCols = 1 To 7 Step 1
                    Select Case nCols
                        Case 1
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "월"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "2"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 2
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "화"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "3"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 3
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "수"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "4"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 4
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "목"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "5"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 5
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "금"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "6"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 6
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "토"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "7"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 7
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "일"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "1"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                    End Select
                Next nCols
                        
                    
                '<< 시간표 테이블 >>
                    sprTimeTable.Col = SpreadHeader:        sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = 6
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 2:    sprTimeTable.ColWidth(sprTimeTable.Col) = 4
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("BASE_CLASS")) = False Then sTmp = Trim(.Fields("BASE_CLASS"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 3:    sprTimeTable.ColWidth(sprTimeTable.Col) = 6
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("DAMIM")) = False Then sTmp = Trim(.Fields("DAMIM"))
                        sprTimeTable.Text = sTmp
                
                
                '<< 요일 만들기 >>
                sprTimeTable.MaxCols = 70
                For nCols = 1 To 7 Step 1
                    Select Case nCols
                        Case 1
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "월"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "2"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 2
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "화"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "3"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 3
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "수"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "4"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 4
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "목"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "5"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 5
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "금"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "6"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 6
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "토"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "7"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 7
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "일"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "1"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                    End Select
                Next nCols
                
                .MoveNext
                
            Next nRec
            
            
            '>> 구분선 긋기
            For nRow = 1 To sprWork.MaxRows Step 1
                For nCol = 1 To sprWork.MaxCols Step 1
                    sprWork.Row = nRow
                    sprWork.Col = nCol
                    
                    If (nCol Mod 10) = 0 Then
                        sprWork.SetCellBorder sprWork.Col, sprWork.Row, sprWork.Col, sprWork.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    End If
                Next nCol
                
                sprWork.SetCellBorder 1, sprWork.Row, sprWork.MaxCols, sprWork.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
            Next nRow
            
            For nRow = 1 To sprTimeTable.MaxRows Step 1
                For nCol = 1 To sprTimeTable.MaxCols Step 1
                    sprTimeTable.Row = nRow
                    sprTimeTable.Col = nCol
                    
                    If (nCol Mod 10) = 0 Then
                        sprTimeTable.SetCellBorder sprTimeTable.Col, sprTimeTable.Row, sprTimeTable.Col, sprTimeTable.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    End If
                Next nCol
                
                sprTimeTable.SetCellBorder 1, sprTimeTable.Row, sprTimeTable.MaxCols, sprTimeTable.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
            Next nRow
            
        End If
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    
    Exit Sub
ErrStmt:
    
    MsgBox "반 조회중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 구성"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
    
End Sub

















'<< 강사/과목별 시수조회
Private Sub cmdTotSisu_Click()
    Dim nLsnCount       As Long
    
    On Error GoTo ErrStmt
    
    sprWork.ColHeaderRows = 1
    sprWork.RowHeaderCols = 1
    
    sprTimeTable.ColHeaderRows = 1
    sprTimeTable.RowHeaderCols = 1
    
    sprTotSisu.MaxRows = 0
    sprLsnSisu.MaxRows = 0
    
    nLsnCount = Find_LsnCount               '< 반 count
    
    If nLsnCount > 0 Then
        Call Construct_Base_TimeTable(nLsnCount)
        
        cmdTotSisu.Tag = "FIND"
            Call Find_Sisu_TotalData
            Call cmdShowTimeTable_Click     '< 전체 시간표 조회
        cmdTotSisu.Tag = ""
        
    End If
    
    MsgBox "시간표 조회하였습니다.", vbInformation + vbOKOnly, "시간표 조회"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    
End Sub

Private Sub Find_Sisu_TotalData()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    
    
    On Error GoTo ErrStmt
    
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, SISUCD, TCR_CL, TCRGBN, TCRNM, SUBJNM, TOT_SISU, TMR_SISU"
    sStr = sStr & "    FROM ("
    sStr = sStr & "          SELECT ACID, SISUCD, MAX(TCR_CL) AS TCR_CL, MAX(TCRGBN) AS TCRGBN, MAX(TCRNM) AS TCRNM, MAX(SUBJNM) AS SUBJNM,"
    sStr = sStr & "                 SUM(NVL(TOT_SISU,0)) AS TOT_SISU, SUM(NVL(TMR_SISU,0)) AS TMR_SISU"
    sStr = sStr & "            FROM ("
    sStr = sStr & "                  /* 전체 시간표 시수로 등록한 내용 */"
    sStr = sStr & "                  SELECT ACID, SISUCD, 0 AS TCR_CL, '' AS TCRGBN, '' AS TCRNM, '' AS SUBJNM, 0 AS TOT_SISU, SUM(NVL(SISU,0)) AS TMR_SISU"
    sStr = sStr & "                    FROM SDTRX50TB"
    sStr = sStr & "                   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                     AND LSNCD"
    sStr = sStr & "                      IN (SELECT LSNCD"
    sStr = sStr & "                            FROM SDLSN01TB"
    sStr = sStr & "                           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                          )"
    sStr = sStr & "                   GROUP BY ACID, SISUCD"
    sStr = sStr & "                  UNION ALL"
    sStr = sStr & "                  /* 강사.과목별로 지정한 총 시수 */"
    sStr = sStr & "                  SELECT ACID, SISUCD, MAX(TCR_CL) AS TCR_CL, TCRGBN, TCRNM, SUBJNM, SUM(NVL(SISU,0)) AS TOT_SISU, 0 AS TMR_SISU"
    sStr = sStr & "                    FROM (SELECT A.ACID, A.SISUCD, A.TCR_CL, A.TCRNM, A.SUBJNM, A.TCRGBN, "
    sStr = sStr & "                                 B.LSNCD, B.SISU"
    sStr = sStr & "                            FROM SDTCR01TB A, SDTCR11TB B"
    sStr = sStr & "                           WHERE A.ACID   = B.ACID"
    sStr = sStr & "                             AND A.SISUCD = B.SISUCD"
    sStr = sStr & "                             AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                             AND B.LSNCD"
    sStr = sStr & "                              IN (SELECT LSNCD"
    sStr = sStr & "                                    FROM SDLSN01TB"
    sStr = sStr & "                                   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                  )"
    sStr = sStr & "                          )"
    sStr = sStr & "                   GROUP BY ACID, SISUCD, TCRNM, SUBJNM, TCRGBN"
    sStr = sStr & "                  )"
    sStr = sStr & "           GROUP BY ACID, SISUCD"
    sStr = sStr & "          )"
    sStr = sStr & "   ORDER BY TCRNM, SUBJNM "
    
    
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
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
                sprTotSisu.MaxRows = sprTotSisu.MaxRows + 1
                sprTotSisu.Row = sprTotSisu.MaxRows:                sprTotSisu.RowHeight(sprTotSisu.Row) = nRowHeight
                
                sprTotSisu.Col = 1:                         sTmp = " "
                    If IsNull(.Fields("ACID")) = False Then
                        sTmp = Trim(.Fields("ACID"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("SISUCD")) = False Then
                        sTmp = Trim(.Fields("SISUCD"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCRGBN")) = False Then
                        sTmp = Trim(.Fields("TCRGBN"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCR_CL")) = False Then
                        sTmp = Trim(.Fields("TCR_CL"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCR_CL")) = False Then
                        sprTotSisu.BackColor = CLng(.Fields("TCR_CL"))
                        sprTotSisu.BackColorStyle = BackColorStyleUnderGrid
                    End If
                
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCRNM")) = False Then
                        sTmp = Trim(.Fields("TCRNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("SUBJNM")) = False Then
                        sTmp = Trim(.Fields("SUBJNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprTotSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        nTmp = 0
                    If IsNumeric(.Fields("TOT_SISU")) = True Then
                        nTmp = CDbl(.Fields("TOT_SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprTotSisu, 0, -99999, 99999, "", nTmp)
                sprTotSisu.Col = sprTotSisu.Col + 1:        nTmp = 0
                    If IsNumeric(.Fields("TMR_SISU")) = True Then
                        nTmp = CDbl(.Fields("TMR_SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprTotSisu, 0, -99999, 99999, "", nTmp)
                    
                sprTotSisu.Col = sprTotSisu.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprTotSisu)
                    sprTotSisu.Value = 0
                    
                .MoveNext
                
            Next nRec
        End If
        
        sprTotSisu.Row = 1:       sprTotSisu.Row2 = sprTotSisu.MaxRows
        sprTotSisu.Col = 1:       sprTotSisu.Col2 = 4
        sprTotSisu.BlockMode = True
            sprTotSisu.BackColor = basModule.BackColor1
            sprTotSisu.BackColorStyle = BackColorStyleUnderGrid
        sprTotSisu.BlockMode = False
        
        sprTotSisu.Row = 1:       sprTotSisu.Row2 = sprTotSisu.MaxRows
        sprTotSisu.Col = 6:       sprTotSisu.Col2 = sprTotSisu.MaxCols
        sprTotSisu.BlockMode = True
            sprTotSisu.BackColor = basModule.BackColor1
            sprTotSisu.BackColorStyle = BackColorStyleUnderGrid
        sprTotSisu.BlockMode = False

    '>> spread lock
        sprTotSisu.Row = 1:       sprTotSisu.Row2 = sprTotSisu.MaxRows
        sprTotSisu.Col = 1:       sprTotSisu.Col2 = sprTotSisu.MaxCols
        sprTotSisu.BlockMode = True
            sprTotSisu.Lock = True
            sprTotSisu.Protect = True
        sprTotSisu.BlockMode = False
    End With
    
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    If cmdTotSisu.Tag = "" Then
        MsgBox "시수 조회하였습니다.", vbInformation + vbOKOnly, "시수조회"
    End If
    
    Exit Sub
ErrStmt:
    
    MsgBox "전체 강사별 시수조회중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시수조회"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
End Sub








'## 선택
Private Sub sprTotSisu_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sSchCD      As String
    Dim sSisuCD     As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprTotSisu
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 6:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:        .Value = 0
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 6:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row
        .Col = .MaxCols:        .Value = 1
        
        .Tag = Trim(CStr(Row))
        fpsprTotSisu_Row.Value = Row
        
        .Row = Row
        .Col = 1:       sSchCD = Trim(.Text):       txtSelSchCD.Text = sSchCD
        .Col = 2:       sSisuCD = Trim(.Text):      txtSelSisuCD.Text = sSisuCD
        
        
        Call Det_Lsn_Sisu_Data(sSchCD, sSisuCD)
        
        
    End With
End Sub



'## 색코드 처리
Private Sub sprTotSisu_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim nColor      As Long
    
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBParam     As ADODB.Parameter

    Dim sTmp        As String
    Dim nTmp        As Long

    Dim sStr        As String
    Dim nExe        As Long
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
        
        sprTotSisu.Row = Row
        sprTotSisu.Col = 1:     sSchCD = Trim(sprTotSisu.Text)
        sprTotSisu.Col = 2:     sSisuCD = Trim(sprTotSisu.Text)
        
        
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
    
        nExe = 0
        DBCmd.Execute nExe, , -1
    
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
    
        If nExe = 1 Then
            basDataBase.DBConn.CommitTrans
            
            With sprTotSisu
                .Row = Row
                .Col = Col - 1
                    sTmp = sSisuCD
                    Call basFunction.Set_SprType_Text(sprTotSisu, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
                .Col = Col
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = nColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
            End With
            
            MsgBox "색상을 등록하였습니다.", vbInformation + vbOKOnly, "색상 선택하기"
            
        Else
            basDataBase.DBConn.RollbackTrans
            
            With sprTotSisu
                .Row = Row
                .Col = Col - 1
                    Call basFunction.Set_SprType_Text(sprTotSisu, "center", "left", 1, "")
                    
                .Col = Col
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
            End With
            
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






'## 선택시 해당 반의 시수내역 조회
Private Sub Det_Lsn_Sisu_Data(ByVal aSchCD As String, ByVal aSisuCD As String)
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    
    sprLsnSisu.MaxRows = 0
    On Error GoTo ErrStmt
    
    
    sStr = ""
    sStr = sStr & "  SELECT A.ACID, A.TCRGBN, A.TCRNM, A.SUBJNM, A.LSNCD, GET_LSNNM(A.LSNCD) AS LSNNM, TCR_CL, "
    sStr = sStr & "         NVL(A.LSN_SISU,0) AS LSN_SISU, NVL(B.SEL_SISU,0) AS SEL_SISU"
    sStr = sStr & "    FROM (SELECT A.ACID, A.TCRGBN, A.TCRNM, A.SUBJNM, MAX(A.TCR_CL) AS TCR_CL, "
    sStr = sStr & "                 B.LSNCD, SUM(NVL(B.SISU,0)) AS LSN_SISU"
    sStr = sStr & "            FROM SDTCR01TB A, SDTCR11TB B"
    sStr = sStr & "           WHERE A.ACID   = B.ACID"
    sStr = sStr & "             AND A.SISUCD = B.SISUCD"
    sStr = sStr & "             AND A.ACID   = '" & aSchCD & "'"
    sStr = sStr & "             AND A.SISUCD = " & aSisuCD
    sStr = sStr & "           GROUP BY A.ACID, A.TCRGBN, A.TCRNM, A.SUBJNM, B.LSNCD"
    sStr = sStr & "          ) A,"
    sStr = sStr & "         (SELECT ACID, LSNCD, SUM(NVL(SISU,0)) AS SEL_SISU"
    sStr = sStr & "            FROM SDTRX50TB"
    sStr = sStr & "           WHERE ACID   = '" & aSchCD & "'"
    sStr = sStr & "             AND SISUCD = " & aSisuCD
    sStr = sStr & "           GROUP BY ACID, LSNCD"
    sStr = sStr & "          ) B"
    sStr = sStr & "   WHERE A.ACID  = B.ACID (+)"
    sStr = sStr & "     AND A.LSNCD = B.LSNCD (+)"
    
    
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
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprLsnSisu.MaxRows = sprLsnSisu.MaxRows + 1
                sprLsnSisu.Row = sprLsnSisu.MaxRows:                sprLsnSisu.RowHeight(sprLsnSisu.Row) = nRowHeight
                
                sprLsnSisu.Col = 1:                         sTmp = " "
                    If IsNull(.Fields("ACID")) = False Then
                        sTmp = Trim(.Fields("ACID"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCRGBN")) = False Then
                        sTmp = Trim(.Fields("TCRGBN"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCR_CL")) = False Then
                        sTmp = Trim(.Fields("TCR_CL"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCR_CL")) = False Then
                        sprLsnSisu.BackColor = CLng(.Fields("TCR_CL"))
                        sprLsnSisu.BackColorStyle = BackColorStyleUnderGrid
                    End If
                
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("TCRNM")) = False Then
                        sTmp = Trim(.Fields("TCRNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("SUBJNM")) = False Then
                        sTmp = Trim(.Fields("SUBJNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("LSNCD")) = False Then
                        sTmp = Trim(.Fields("LSNCD"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        sTmp = " "
                    If IsNull(.Fields("LSNNM")) = False Then
                        sTmp = Trim(.Fields("LSNNM"))
                    End If
                    Call basFunction.Set_SprType_Text(sprLsnSisu, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        nTmp = 0
                    If IsNumeric(.Fields("LSN_SISU")) = True Then
                        nTmp = CDbl(.Fields("LSN_SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprLsnSisu, 0, -99999, 99999, "", nTmp)
                sprLsnSisu.Col = sprLsnSisu.Col + 1:        nTmp = 0
                    If IsNumeric(.Fields("SEL_SISU")) = True Then
                        nTmp = CDbl(.Fields("SEL_SISU"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprLsnSisu, 0, -99999, 99999, "", nTmp)
                    
                sprLsnSisu.Col = sprLsnSisu.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprLsnSisu)
                    sprLsnSisu.Value = 0
                    
                .MoveNext
            Next nRec
        End If
        
        sprLsnSisu.Row = 1:       sprLsnSisu.Row2 = sprLsnSisu.MaxRows
        sprLsnSisu.Col = 1:       sprLsnSisu.Col2 = 3
        sprLsnSisu.BlockMode = True
            sprLsnSisu.BackColor = basModule.BackColor1
            sprLsnSisu.BackColorStyle = BackColorStyleUnderGrid
        sprLsnSisu.BlockMode = False
        
        sprLsnSisu.Row = 1:       sprLsnSisu.Row2 = sprLsnSisu.MaxRows
        sprLsnSisu.Col = 5:       sprLsnSisu.Col2 = sprLsnSisu.MaxCols
        sprLsnSisu.BlockMode = True
            sprLsnSisu.BackColor = basModule.BackColor1
            sprLsnSisu.BackColorStyle = BackColorStyleUnderGrid
        sprLsnSisu.BlockMode = False
        

    '>> spread lock
        sprLsnSisu.Row = 1:       sprLsnSisu.Row2 = sprLsnSisu.MaxRows
        sprLsnSisu.Col = 1:       sprLsnSisu.Col2 = sprLsnSisu.MaxCols
        sprLsnSisu.BlockMode = True
            sprLsnSisu.Lock = True
            sprLsnSisu.Protect = True
        sprLsnSisu.BlockMode = False
    End With
    
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    
    MsgBox "반별 시수조회중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "세부 시수조회"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing

End Sub




'>> 1로 선택된 부분을 저장가능상태로 바꾸어 줌.
Private Sub sprWork_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub

    With sprWork
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
            
        ElseIf .Text = "S" Then
            .Text = "1"
            .SetCellBorder .Col, .Row, .Col, .Row, 16, basModule.GridColor2, CellBorderStyleSolid
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = txtnColor.BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
        End If
    End With

End Sub







'## 색상 등록하기
Private Sub sprWork_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sLsnCD      As String
    Dim sPreData    As String
    Dim sData       As String
    Dim sQuestion   As String
    
    Dim sStr        As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim ni          As Integer
    Dim nLsnCount   As Long
    
    If Col = SpreadHeader Then Exit Sub
    If Col = SpreadHeader + 1 Then Exit Sub
    
    On Error GoTo ErrStmt
    
    '> 기본교실 & 담임교사
    If Col = SpreadHeader + 2 Or Col = SpreadHeader + 3 Then
        
        With sprWork
            .Row = Row
            .Col = SpreadHeader
                sLsnCD = Trim(.Text)
        
            sData = ""
            Select Case Col
                Case SpreadHeader + 2
                    
                    .Row = Row
                    .Col = SpreadHeader + 2
                        sPreData = Trim(.Text)
                        
                    
                    If sPreData > " " Then
                        sQuestion = "기존 교실 : " & sPreData & vbCrLf & _
                                    "바꾸고자 하는 교실명을 넣어주세요." & vbCrLf & _
                                    "단, 삭제를 하실경우 - (하이픈)을 넣어주세요."
                    Else
                        sQuestion = "교실명을 넣으세요."
                    End If
                        
                    sData = ""
                    sData = InputBox(sQuestion, "기본교실 넣기", "")
                
                    If sData = "" Then Exit Sub
                    
                    sStr = ""
                    sStr = sStr & "     UPDATE SDLSN01TB "
                    If sData = "-" Then
                        sStr = sStr & "    SET BASE_CLASS = '' "
                    Else
                        sStr = sStr & "    SET BASE_CLASS = '" & sData & "' "
                    End If
                    sStr = sStr & "      WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                    sStr = sStr & "        AND LSNCD  = '" & sLsnCD & "'"
                    
                Case SpreadHeader + 3
                    
                    .Row = Row
                    .Col = SpreadHeader + 3
                        sPreData = Trim(.Text)
                        
                    
                    If sPreData > " " Then
                        sQuestion = "기존 담임 : " & sPreData & vbCrLf & _
                                    "바꾸고자 하는 담임명을 넣어주세요." & _
                                    "단, 삭제를 하실경우 - (하이픈)을 넣어주세요."
                    Else
                        sQuestion = "담임명을 넣으세요."
                    End If
                    
                    sData = ""
                    sData = InputBox(sQuestion, "기본교실 넣기", "")
                    
                    If sData = "" Then Exit Sub
                    
                    sStr = ""
                    sStr = sStr & "     UPDATE SDLSN01TB "
                    If sData = "-" Then
                        sStr = sStr & "    SET DAMIM = '' "
                    Else
                        sStr = sStr & "    SET DAMIM = '" & sData & "' "
                    End If
                    sStr = sStr & "      WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
                    sStr = sStr & "        AND LSNCD = '" & sLsnCD & "'"
                    
            End Select
        
        End With
                
                
        '## 데이터 등록
        basDataBase.DBConn.BeginTrans
    
        Set DBCmd = New ADODB.Command
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        
        nExe = 0
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
                
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
                
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nExe = 1 Then
            basDataBase.DBConn.CommitTrans
                        
            sprWork.ColHeaderRows = 1
            sprWork.RowHeaderCols = 1
            
            sprTimeTable.ColHeaderRows = 1
            sprTimeTable.RowHeaderCols = 1
            
            nLsnCount = Find_LsnCount           '< 반 count
            
            If nLsnCount > 0 Then
                Call Construct_Base_TimeTable(nLsnCount)
                
            End If
            
            MsgBox "등록하였습니다.", vbInformation + vbOKOnly, "반 내역 등록"
        Else
            basDataBase.DBConn.RollbackTrans
            MsgBox "등록중 오류가 발생하였습니다." & vbCrLf & _
                   Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "반 내역 등록"
        End If
            
        Set DBCmd = Nothing
        Set DBParam = Nothing
    End If
         
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    MsgBox "등록중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "반 내역 등록"
           
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
End Sub


'## 시수등록할 반을 선택
Private Sub sprLsnSisu_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sSchCD      As String           '< 학원
    Dim sGbn        As String           '< 인문/자연.. 국영수 사과 (10,20,30    40,50)
    Dim sSelColor   As String           '< 색
    Dim sTeacher    As String           '< 강사
    Dim sGwamok     As String           '< 과목
    Dim sLsnCD      As String           '< 반
    
    Dim nWTotSisu   As Long             '< 반 시수
    Dim nWLsnSisu   As Long             '< 선택시수
    
    Dim nWorkRow    As Long
    Dim nWorkCol    As Long
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprLsnSisu
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 3
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 5:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:        .Value = 0
        
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = 3
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 5:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row
        .Col = .MaxCols:        .Value = 1
        
        .Tag = Trim(CStr(Row))
        fpsprLsnSisuRow.Value = Row
        
        .Row = Row
        .Col = 1:   sSchCD = Trim(.Text)                '< 학원
        .Col = 2:   sGbn = Trim(.Text):     txtSelGbn.Text = sGbn       '< 인문/자연.. 국영수 사과 (10,20,30    40,50)
        .Col = 3:   sSelColor = Trim(.Text)             '< 색
            If sSelColor = "" Then
                txtnColor.BackColor = basModule.WhiteColor
            Else
                txtnColor.BackColor = CLng(sSelColor)
            End If
        .Col = 5:   sTeacher = Trim(.Text)              '< 강사
        .Col = 6:   sGwamok = Trim(.Text)               '< 과목
        .Col = 7:   sLsnCD = Trim(.Text):   txtSelLsnCD.Text = sLsnCD       '< 반
    
        .Col = 9:   nWTotSisu = .Value                  '< 반 시수
        .Col = 10:  nWLsnSisu = .Value                  '< 선택시수
        
        fpWorkSisu.Value = 0
        If nWTotSisu - nWLsnSisu <= 0 Then
        
            With sprWork
                '## [1] 초기화 ##########################################
                For nWorkRow = 1 To .MaxRows Step 1
                    .Row = nWorkRow
                    For nWorkCol = 1 To .MaxCols Step 1
                        .Col = nWorkCol
                            Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                    Next nWorkCol
                Next nWorkRow
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
                
        
            lblStatus.Caption = "선택가능한 시수가 없습니다."
        Else
            
            fpWorkSisu.Value = nWTotSisu - nWLsnSisu    '< 작업가능 시수
        
            Select Case sGbn
                Case "10", "20", "30"       '< 언,수,외
                    Call WorkTable_Schdule_Checks_KME(sSchCD, sGbn, sSelColor, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                    
                    
                Case "40", "50"             '< 사,과
                    Call WorkTable_Schdule_Checks_Tamgu(sSchCD, sGbn, sSelColor, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                    
            End Select
        End If
        
    End With
End Sub





'## 언.수.외 선택인 경우 #############################################################################################################
'## 아래의 작업진행
Private Sub WorkTable_Schdule_Checks_KME(ByVal aSchCD As String, _
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
        
        
        '## [1] 초기화 ##########################################
        For nWorkRow = 1 To .MaxRows Step 1
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
        Next nWorkRow
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        
        '## [2] 작업진행 ########################################
        For nWorkRow = 1 To .MaxRows Step 1
            .Row = nWorkRow
            .Col = SpreadHeader
            
            '## 해당반이 있는 경우
            If StrComp(aLsnCD, Trim(.Text), vbTextCompare) = 0 Then
                
                '> 1. 전체 선택 가능상태 ---------------------------------------------------------------------------------------------------------------
                .Row = nWorkRow
                For nWorkCol = 1 To .MaxCols Step 1
                    .Col = nWorkCol
                        Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "1")
                Next nWorkCol
                
                .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
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
                sStr = sStr & "   WHERE A.ACID   = B.ACID"
                sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
                sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                            '< 2007.12.18 : 계열추가
                sStr = sStr & "     AND A.ACID   = '" & aSchCD & "'"
                sStr = sStr & "     AND A.TRXCD  LIKE (SELECT LSNTYPE||'%'"
                sStr = sStr & "                         FROM SDLSN01TB"
                sStr = sStr & "                        WHERE ACID  = '" & aSchCD & "'"
                sStr = sStr & "                          AND LSNCD = '" & aLsnCD & "'"
                sStr = sStr & "                       ) "
                sStr = sStr & "     AND A.KAEYOL IN   (SELECT KAEYOL"                   '< 2007.12.18 : 계열추가
                sStr = sStr & "                          FROM SDLSN01TB"
                sStr = sStr & "                         WHERE ACID  = '" & aSchCD & "'"
                sStr = sStr & "                           AND LSNCD = '" & aLsnCD & "'"
                sStr = sStr & "                        ) "
                
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
                                    .Col = 1 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:       .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "3"
                                    .Col = 11 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "4"
                                    .Col = 21 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "5"
                                    .Col = 31 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "6"
                                    .Col = 41 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "7"
                                    .Col = 51 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "1"
                                    .Col = 61 + CLng(sLesson) - 1
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
                                    .Col = 1 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:       .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "3"
                                    .Col = 11 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "4"
                                    .Col = 21 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "5"
                                    .Col = 31 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "6"
                                    .Col = 41 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "7"
                                    .Col = 51 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "1"
                                    .Col = 61 + CLng(sLesson) - 1
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
                                    .Col = 1 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:       .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "3"
                                    .Col = 11 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "4"
                                    .Col = 21 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "5"
                                    .Col = 31 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "6"
                                    .Col = 41 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "7"
                                    .Col = 51 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "1"
                                    .Col = 61 + CLng(sLesson) - 1
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
                
                
            End If
        Next nWorkRow
    End With
    
    
    If bChk = False Then
        '> 처리 오류이므로 원상복귀
        With sprWork
            For nWorkRow = 1 To .MaxRows
                .Row = nWorkRow
                For nWorkCol = 1 To .MaxCols Step 1
                    .Col = nWorkCol
                        Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                Next nWorkCol
            Next nWorkRow
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
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
        For nWorkRow = 1 To .MaxRows
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
        Next nWorkRow
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
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
Private Sub WorkTable_Schdule_Checks_Tamgu(ByVal aSchCD As String, _
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
        
        
        '## [1] 초기화 ##########################################
        For nWorkRow = 1 To .MaxRows Step 1
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
        Next nWorkRow
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        
        '## [2] 작업진행 ########################################
        For nWorkRow = 1 To .MaxRows Step 1
            .Row = nWorkRow
            .Col = SpreadHeader
            
            '## 해당반이 있는 경우
            If StrComp(aLsnCD, Trim(.Text), vbTextCompare) = 0 Then
                
                
                
                '> 1. 선택가능 내용 검색 << 사과탐 부분 >> -------------------------------------------------------------------------------------------
                sStr = ""
                sStr = sStr & "  SELECT LESSON, WEEKS"
                sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
                sStr = sStr & "   WHERE A.ACID   = B.ACID"
                sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
                sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"                            '< 2007.12.18 : 계열
                sStr = sStr & "     AND A.ACID   = '" & aSchCD & "'"
                sStr = sStr & "     AND A.TRXCD  LIKE (SELECT LSNTYPE||'%'"
                sStr = sStr & "                          FROM SDLSN01TB"
                sStr = sStr & "                         WHERE ACID  = '" & aSchCD & "'"
                sStr = sStr & "                           AND LSNCD = '" & aLsnCD & "'"
                sStr = sStr & "                        ) "
                sStr = sStr & "     AND A.TRXCD  IN   (SELECT KAEYOL"                   '< 2007.12.18 : 계열
                sStr = sStr & "                          FROM SDLSN01TB"
                sStr = sStr & "                         WHERE ACID  = '" & aSchCD & "'"
                sStr = sStr & "                           AND LSNCD = '" & aLsnCD & "'"
                sStr = sStr & "                        ) "
                
                
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
                                    .Col = 1 + CLng(sLesson) - 1
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
                                    .Col = 11 + CLng(sLesson) - 1
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
                                    
                                Case "5"
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
                                    
                                Case "6"
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
                                    
                                Case "7"
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
                                    
                                Case "1"
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
                                    .Col = 1 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:       .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "3"
                                    .Col = 11 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "4"
                                    .Col = 21 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "5"
                                    .Col = 31 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "6"
                                    .Col = 41 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "7"
                                    .Col = 51 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "1"
                                    .Col = 61 + CLng(sLesson) - 1
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
                                    .Col = 1 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:       .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "3"
                                    .Col = 11 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                        
                                Case "4"
                                    .Col = 21 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "5"
                                    .Col = 31 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "6"
                                    .Col = 41 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "7"
                                    .Col = 51 + CLng(sLesson) - 1
                                        Call basFunction.Set_SprType_Text(sprWork, "CENTER", "CENTER", 1, "")
                                        
                                        .Row2 = .Row:        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    
                                Case "1"
                                    .Col = 61 + CLng(sLesson) - 1
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
                
                
            End If
        Next nWorkRow
    End With
    
    
    If bChk = False Then
        '> 처리 오류이므로 원상복귀
        With sprWork
            For nWorkRow = 1 To .MaxRows
                .Row = nWorkRow
                For nWorkCol = 1 To .MaxCols Step 1
                    .Col = nWorkCol
                        Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
                Next nWorkCol
            Next nWorkRow
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
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
        For nWorkRow = 1 To .MaxRows
            .Row = nWorkRow
            For nWorkCol = 1 To .MaxCols Step 1
                .Col = nWorkCol
                    Call basFunction.Set_SprType_Text(sprWork, "center", "center", 1, "")
            Next nWorkCol
        Next nWorkRow
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
                
    MsgBox "작업 시간표 처리시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "작업 시간표 처리"
    

End Sub
















'>> 시간표
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
    Dim nExe        As Integer
    Dim nAccExe     As Long
    Dim nTotExe     As Long
    
    ReDim uWorkTimeTable(0) As tWorkTimeTable           '< 등록할 자료
    
    On Error GoTo ErrStmt
    
    With sprWork
        nCountChk_S = 0     '< S로 체크되어진 갯수
        
        For nRow_Work = 1 To .MaxRows Step 1
            For nCol_Work = 1 To .MaxCols Step 1
                .Row = nRow_Work
                .Col = nCol_Work
                
                If StrComp(Trim(.Text), "S", vbTextCompare) = 0 Then
                    nCountChk_S = nCountChk_S + 1
                    
                    ReDim Preserve uWorkTimeTable(nCountChk_S) As tWorkTimeTable
                    
                    '## 등록할 데이터 ----------------------------------------------------------------
                    uWorkTimeTable(nCountChk_S).ACID = Trim(txtSelSchCD.Text)       '< 학원
                    .Row = nRow_Work
                        .Col = SpreadHeader
                            uWorkTimeTable(nCountChk_S).LSNCD = Trim(.Text)         '< 반
                    .Row = SpreadHeader + 2
                        .Col = nCol_Work
                            uWorkTimeTable(nCountChk_S).LESSON = Trim(.Text)        '< 교시
                    .Row = SpreadHeader + 1
                        .Col = nCol_Work
                            uWorkTimeTable(nCountChk_S).WEEK = Trim(.Text)          '< 요일
                    uWorkTimeTable(nCountChk_S).SISUCD = Trim(txtSelSisuCD.Text)    '< 시수코드
                    uWorkTimeTable(nCountChk_S).SISU = "1"
                    uWorkTimeTable(nCountChk_S).TRX_CL = Trim(txtnColor.BackColor)  '< 색
                    '---------------------------------------------------------------------------------
                    
                    .SetCellBorder .Col, .Row, .Col, .Row, 16, basModule.GridColor2, CellBorderStyleSolid
                    
                End If
            Next nCol_Work
        Next nRow_Work
    End With

    If fpWorkSisu < nCountChk_S Then     '< 선택가능 시수보다 작아야 합니다.
        MsgBox "현재 선택가능한 시수보다 많습니다." & vbCrLf & _
               "선택가능한 시수는 총 " & Trim(CStr(fpWorkSisu.Value)) & "입니다.", vbExclamation + vbOKOnly, "시간표 등록"
        Exit Sub
    End If
    
    If UBound(uWorkTimeTable) = 0 Then  '< S 로 선택된 내용이 없습니다.
        MsgBox "등록할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 등록"
        Exit Sub
    End If
    
    
    
    
    nExe = 0
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
    
        DBCmd.Execute nExe, , -1
                
                
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nExe = 1 Then
            nAccExe = nAccExe + 1
        End If
        
    Next nRec
    
    If nTotExe = nAccExe Then
        basDataBase.DBConn.CommitTrans
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    
    
    '## 전부 다시 조회 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '       sprTotSisu
    '       sprLsnSisu
    '       sprWork
    '       sprTimeTable
        cmdTotSisu.Tag = "REVIEW"
            sprTotSisu.MaxRows = 0
            Call Find_Sisu_TotalData        ' 계열시수내역 조회
            
            With sprTotSisu
                .Row = fpsprTotSisu_Row.Value:  .Row2 = .Row
                .Col = 1:                       .Col2 = 4
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = fpsprTotSisu_Row.Value:  .Row2 = .Row
                .Col = 6:                       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = .MaxCols:        .Value = 1
            End With
        cmdTotSisu.Tag = ""
        Call cmdFindLsnSisu_Click       ' 해당 시수코드의 내용 조회
        Call cmdFindWork_Click          ' sprLsnSisu 부분의 선택 및 sprWork 에 대한 부분 조회
        Call cmdShowTimeTable_Click     ' sprTimeTable (전체 시간표 조회)
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
































'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<< 시간표 등록후 다시 등록 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'## sprLsnSisu 다시 조회
Private Sub cmdFindLsnSisu_Click()

    If Trim(txtSelSchCD.Text) = "" Then Exit Sub
    If Trim(txtSelSisuCD.Text) = "" Then Exit Sub

    Call Det_Lsn_Sisu_Data(Trim(txtSelSchCD.Text), Trim(txtSelSisuCD.Text))
    
End Sub

'## sprWork 내역 조회
Private Sub cmdFindWork_Click()
    
    Dim nLsnCount   As Long
    
    Dim sSchCD      As String           '< 학원
    Dim sGbn        As String           '< 인문/자연.. 국영수 사과 (10,20,30    40,50)
    Dim sSelColor   As String           '< 색
    Dim sTeacher    As String           '< 강사
    Dim sGwamok     As String           '< 과목
    Dim sLsnCD      As String           '< 반
    
    
    Dim nWTotSisu   As Long             '< 반 시수
    Dim nWLsnSisu   As Long             '< 선택시수
    
    
    sprWork.MaxRows = 0
    sprWork.MaxCols = 0
    sprWork.ColHeaderRows = 1
    sprWork.RowHeaderCols = 1
    
    sprTimeTable.MaxRows = 0
    sprTimeTable.MaxCols = 0
    sprTimeTable.ColHeaderRows = 1
    sprTimeTable.RowHeaderCols = 1
        
    nLsnCount = Find_LsnCount           '< 반 count
    If nLsnCount > 0 Then
        Call Construct_init_sprWork(nLsnCount)
                
    End If
    
    If sprLsnSisu.MaxRows = 0 Then Exit Sub
    
    
    With sprLsnSisu
        
        .Row = fpsprLsnSisuRow.Value:   .Row2 = .Row
        .Col = 1:                       .Col2 = 3
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = fpsprLsnSisuRow.Value:   .Row2 = .Row
        .Col = 5:                       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = fpsprLsnSisuRow.Value
        .Col = .MaxCols:        .Value = 1
        
        .Tag = Trim(CStr(fpsprLsnSisuRow.Value))
        fpsprLsnSisuRow.Value = fpsprLsnSisuRow.Value
        
        .Row = fpsprLsnSisuRow.Value
        sSchCD = Trim(txtSelSchCD.Text)                 '< 학원
        sGbn = Trim(txtSelGbn)                          '< 인문/자연.. 국영수 사과 (10,20,30    40,50)
        sSelColor = Trim(txtnColor.BackColor)           '< 색
            If sSelColor = "" Then
                txtnColor.BackColor = basModule.WhiteColor
            Else
                txtnColor.BackColor = CLng(sSelColor)
            End If
            
        .Col = 5:   sTeacher = Trim(.Text)              '< 강사
        .Col = 6:   sGwamok = Trim(.Text)               '< 과목
        
        sLsnCD = Trim(txtSelLsnCD.Text)                 '< 반
    
        .Col = 9:   nWTotSisu = .Value                  '< 반 시수
        .Col = 10:  nWLsnSisu = .Value                  '< 선택시수
        
        fpWorkSisu.Value = 0
        If nWTotSisu - nWLsnSisu <= 0 Then
            lblStatus.Caption = "선택가능한 시수가 없습니다."
        Else
            
            fpWorkSisu.Value = nWTotSisu - nWLsnSisu    '< 작업가능 시수
        
            Select Case sGbn
                Case "10", "20", "30"       '< 언,수,외
                    Call WorkTable_Schdule_Checks_KME(sSchCD, sGbn, sSelColor, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                    
                    
                Case "40", "50"             '< 사,과
                    Call WorkTable_Schdule_Checks_Tamgu(sSchCD, sGbn, sSelColor, sTeacher, sGwamok, sLsnCD, nWTotSisu, nWLsnSisu)
                    
            End Select
        End If
        
    End With
    
End Sub


'=====================================================================================================
' sprWork의 내용을 다시 초기화
'=====================================================================================================
Private Sub Construct_init_sprWork(ByVal aLsnCount As Long)
    
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCols       As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    '/* cols & rows 조정 */
    If optView(0).Value = True Then
        nTtRowHeight = 25
        nTtColWidth = 6
    ElseIf optView(1).Value = True Then
        nTtRowHeight = 20
        nTtColWidth = 2
    End If
    
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, BASE_CLASS, DAMIM"
    sStr = sStr & "    From SDLSN01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    
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
        If .RecordCount > 0 Then
            .MoveFirst
            
            sprWork.Col = SpreadHeader:         sprWork.ColHidden = False
            sprWork.MaxRows = aLsnCount
            sprWork.RowHeaderCols = 4
            sprWork.MaxCols = 0
            sprWork.ColHeaderRows = 3
            sprWork.Col = SpreadHeader:         sprWork.ColHidden = True
            sprWork.Row = SpreadHeader + 1:     sprWork.RowHidden = True
            
            
            For nRec = 1 To .RecordCount Step 1
            
                '<< 작업테이블 >>
                    sprWork.Col = SpreadHeader:             sprWork.ColWidth(sprWork.Col) = nTtColWidth
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 1:         sprWork.ColWidth(sprWork.Col) = 6
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 2:         sprWork.ColWidth(sprWork.Col) = 4
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("BASE_CLASS")) = False Then sTmp = Trim(.Fields("BASE_CLASS"))
                        sprWork.Text = sTmp
                    
                    sprWork.Col = SpreadHeader + 3:         sprWork.ColWidth(sprWork.Col) = 6
                    sprWork.Row = nRec:                     sprWork.RowHeight(sprWork.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("DAMIM")) = False Then sTmp = Trim(.Fields("DAMIM"))
                        sprWork.Text = sTmp
                        
                '<< 요일 만들기 >>
                sprWork.MaxCols = 70
                For nCols = 1 To 7 Step 1
                    Select Case nCols
                        Case 1
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "월"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "2"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 2
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "화"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "3"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 3
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "수"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "4"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 4
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "목"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "5"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 5
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "금"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "6"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 6
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "토"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "7"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                        Case 7
                            sprWork.Col = (nCols - 1) * 10 + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                sprWork.Row = SpreadHeader:         sprWork.Text = "일"
                                sprWork.AddCellSpan sprWork.Col, sprWork.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprWork.Row = SpreadHeader + 1:     sprWork.Text = "1"
                                    sprWork.Row = SpreadHeader + 2:     sprWork.Text = Trim(CStr(nTmp))
                                    
                                    sprWork.Col = sprWork.Col + 1:    sprWork.ColWidth(sprWork.Col) = nTtColWidth
                                Next nTmp
                    End Select
                Next nCols
                
                .MoveNext
                
            Next nRec
        End If
    End With
    
    '>> 구분선 긋기
    For nRow = 1 To sprWork.MaxRows Step 1
        For nCol = 1 To sprWork.MaxCols Step 1
            sprWork.Row = nRow
            sprWork.Col = nCol
            
            If (nCol Mod 10) = 0 Then
                sprWork.SetCellBorder sprWork.Col, sprWork.Row, sprWork.Col, sprWork.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            End If
        Next nCol
        
        sprWork.SetCellBorder 1, sprWork.Row, sprWork.MaxCols, sprWork.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
    Next nRow
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    
    Exit Sub
ErrStmt:
    
    MsgBox "반 조회중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 구성"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
End Sub








'## 전체 시간표 조회
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
    
    Dim sLsnCD          As String
    Dim sLesson         As String
    Dim sWeeks          As String
    
    Dim sTcrNM          As String
    Dim sSubjNM         As String
    Dim sTcr_CL         As String
    Dim sDisp_Text      As String
    
    Dim nTimeTableRow   As Long
    
    On Error GoTo ErrStmt
    
    
    
    If cmdTotSisu.Tag = "FIND" Then
        'no action
        
    Else
        '-- 시간표의 기본 반을 보여주는 부분 -------------------------------
        sprTimeTable.ColHeaderRows = 1
        sprTimeTable.RowHeaderCols = 1
        
        nLsnCount = Find_LsnCount           '< 반 count
        If nLsnCount > 0 Then
            Call Construct_init_sprTimeTable(nLsnCount)
        End If
        '--------------------------------------------------------------------
    End If
    
    '## 전체내역 모두 조회
    sStr = ""
    sStr = sStr & "  SELECT LSNCD,"
    sStr = sStr & "         LESSON, WEEKS,"
    sStr = sStr & "         GET_TCRNM_TCR01(ACID,SISUCD) AS TCRNM,"
    sStr = sStr & "         GET_SUBJNM_TCR01(ACID,SISUCD) AS SUBJNM,"
    sStr = sStr & "         SISU,"
    sStr = sStr & "         TRX_CL"
    sStr = sStr & "    FROM SDTRX50TB"
    sStr = sStr & "   WHERE ACID = ? "
    
    
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
                
'>> 분원
    sTmp = Trim(basModule.SchCD)
    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
                
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sprTimeTable.Row = SpreadHeader
        sprTimeTable.Col = SpreadHeader:     sprTimeTable.Text = "반코드":    sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
        sprTimeTable.Col = SpreadHeader + 1: sprTimeTable.Text = "반":        sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
        sprTimeTable.Col = SpreadHeader + 2: sprTimeTable.Text = "기본반":    sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
        sprTimeTable.Col = SpreadHeader + 3: sprTimeTable.Text = "담임":      sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 1, 3
    
                
    If DBRec.RecordCount > 0 Then
        DBRec.MoveFirst
        For nRec = 1 To DBRec.RecordCount Step 1
                        
            If IsNull(DBRec.Fields("LESSON")) = False And IsNull(DBRec.Fields("WEEKS")) = False Then
                            
                sLesson = Trim(DBRec.Fields("LESSON"))
                sWeeks = Trim(DBRec.Fields("WEEKS"))
                            
                            
                For nTimeTableRow = 1 To sprTimeTable.MaxRows Step 1
                    sprTimeTable.Row = nTimeTableRow
                    sprTimeTable.Col = SpreadHeader
                    
                        sLsnCD = Trim(sprTimeTable.Text)        '< 반 코드
                        
                    If IsNull(DBRec.Fields("LSNCD")) = False Then
                        
                        If StrComp(sLsnCD, DBRec.Fields("LSNCD"), vbTextCompare) = 0 Then
                        
                            '>> 강사명, 과목명, 색을 지정
                            sTcrNM = "":    If IsNull(DBRec.Fields("TCRNM")) = False Then sTcrNM = Trim(DBRec.Fields("TCRNM"))
                                If optView(1).Value = True Then
                                    sTcrNM = basFunction.MidKor(DBRec.Fields("TCRNM"), 1, 2)
                                End If
                            sSubjNM = "":   If IsNull(DBRec.Fields("SUBJNM")) = False Then sSubjNM = Trim(DBRec.Fields("SUBJNM"))
                                If optView(1).Value = True Then
                                    sSubjNM = basFunction.MidKor(DBRec.Fields("SUBJNM"), 1, 2)
                                End If
                            sTcr_CL = "":   If IsNull(DBRec.Fields("TRX_CL")) = False Then sTcr_CL = Trim(DBRec.Fields("TRX_CL"))
                            
                            sDisp_Text = sSubjNM & vbCrLf & sTcrNM  '< spread cell 에 보여질 내용
                            
                            
                            
                            sprTimeTable.Row = nTimeTableRow        '< 현재 ROW
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
                    End If
                Next nTimeTableRow
            End If
            
            DBRec.MoveNext
        Next nRec
    End If
    
    With sprTimeTable
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "전체시간표 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체 시간표 조회"
    
End Sub





'=====================================================================================================
' sprTimeTable의 내용을 다시 초기화
'=====================================================================================================
Private Sub Construct_init_sprTimeTable(ByVal aLsnCount As Long)
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCols       As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    
    '/* cols & rows 조정 */
    If optView(0).Value = True Then
        nTtRowHeight = 25
        nTtColWidth = 6
    ElseIf optView(1).Value = True Then
        nTtRowHeight = 20
        nTtColWidth = 2
    End If
    
    
    On Error GoTo ErrStmt
    
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, BASE_CLASS, DAMIM"
    sStr = sStr & "    From SDLSN01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    
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
        If .RecordCount > 0 Then
            .MoveFirst
            
            sprTimeTable.Col = SpreadHeader:        sprTimeTable.ColHidden = False
            sprTimeTable.MaxRows = aLsnCount
            sprTimeTable.RowHeaderCols = 4
            sprTimeTable.MaxCols = 0
            sprTimeTable.ColHeaderRows = 3
            
            sprTimeTable.Col = SpreadHeader:    sprTimeTable.ColHidden = True
            sprTimeTable.Row = SpreadHeader + 1:    sprTimeTable.RowHidden = True
            
            For nRec = 1 To .RecordCount Step 1
                
                '<< 시간표 테이블 >>
                    sprTimeTable.Col = SpreadHeader:        sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = 6
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 2:    sprTimeTable.ColWidth(sprTimeTable.Col) = 4
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("BASE_CLASS")) = False Then sTmp = Trim(.Fields("BASE_CLASS"))
                        sprTimeTable.Text = sTmp
                    
                    sprTimeTable.Col = SpreadHeader + 3:    sprTimeTable.ColWidth(sprTimeTable.Col) = 6
                    sprTimeTable.Row = nRec:                sprTimeTable.RowHeight(sprTimeTable.Row) = nTtRowHeight
                        sTmp = " ":  If IsNull(.Fields("DAMIM")) = False Then sTmp = Trim(.Fields("DAMIM"))
                        sprTimeTable.Text = sTmp
                
                
                '<< 요일 만들기 >>
                sprTimeTable.MaxCols = 70
                For nCols = 1 To 7 Step 1
                    Select Case nCols
                        Case 1
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "월"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "2"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 2
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "화"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "3"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 3
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "수"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "4"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 4
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "목"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "5"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 5
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "금"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "6"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 6
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "토"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "7"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                        Case 7
                            sprTimeTable.Col = (nCols - 1) * 10 + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                sprTimeTable.Row = SpreadHeader:         sprTimeTable.Text = "일"
                                sprTimeTable.AddCellSpan sprTimeTable.Col, sprTimeTable.Row, 10, 1
                                
                                '## column은 정해진 상태에서 처리
                                For nTmp = 1 To 10 Step 1
                                    sprTimeTable.Row = SpreadHeader + 1:     sprTimeTable.Text = "1"
                                    sprTimeTable.Row = SpreadHeader + 2:     sprTimeTable.Text = Trim(CStr(nTmp))
                                    
                                    sprTimeTable.Col = sprTimeTable.Col + 1:    sprTimeTable.ColWidth(sprTimeTable.Col) = nTtColWidth
                                Next nTmp
                    End Select
                Next nCols
                
                .MoveNext
                
            Next nRec
            
            '>> 구분선 긋기
            For nRow = 1 To sprTimeTable.MaxRows Step 1
                For nCol = 1 To sprTimeTable.MaxCols Step 1
                    sprTimeTable.Row = nRow
                    sprTimeTable.Col = nCol
                    
                    If (nCol Mod 10) = 0 Then
                        sprTimeTable.SetCellBorder sprTimeTable.Col, sprTimeTable.Row, sprTimeTable.Col, sprTimeTable.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    End If
                Next nCol
                
                sprTimeTable.SetCellBorder 1, sprTimeTable.Row, sprTimeTable.MaxCols, sprTimeTable.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
            Next nRow
            
            
        End If
    End With
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    
    Exit Sub
ErrStmt:
    
    MsgBox "시간표 초기화중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 구성"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
    
End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>








'## 등록된 시간표 내역 삭제
Private Sub cmdDelTimeTable_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    Dim nExe        As Integer
    
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
        
        
        sAcID = Trim(basModule.SchCD)
        .Row = .ActiveRow
        .Col = SpreadHeader:        sLsnCD = Trim(.Text)
        
        .Col = .ActiveCol
        .Row = SpreadHeader + 1:    sWeeks = Trim(.Text)
        .Row = SpreadHeader + 2:    sLesson = Trim(.Text)
        
        
        '## 전체내역 모두 조회
        sStr = ""
        sStr = sStr & "  SELECT GET_TCRNM_TCR01(ACID,SISUCD) AS TCRNM,"
        sStr = sStr & "         GET_SUBJNM_TCR01(ACID,SISUCD) AS SUBJNM "
        sStr = sStr & "    FROM SDTRX50TB"
        sStr = sStr & "   WHERE ACID   = '" & sAcID & "'"
        sStr = sStr & "     AND LSNCD  = '" & sLsnCD & "'"
        sStr = sStr & "     AND LESSON = " & sLesson
        sStr = sStr & "     AND WEEKS  = " & sWeeks
        
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
                
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
                    
                    
        If DBRec.RecordCount <> 1 Then
            MsgBox "처리할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 선택삭제"
            
            Set DBCmd = Nothing
            Set DBRec = Nothing
            
            Exit Sub
        Else
            DBRec.MoveFirst
            
            If IsNull(DBRec.Fields("TCRNM")) = False And IsNull(DBRec.Fields("SUBJNM")) = False Then
                            
                sTcrNM = Trim(DBRec.Fields("TCRNM"))
                sSubjNM = Trim(DBRec.Fields("SUBJNM"))
                
                If MsgBox("과목【 " & sSubjNM & " 】" & vbCrLf & _
                          "강사【 " & sTcrNM & " 】 내용을 삭제하시겠습니까?", vbQuestion + vbYesNo, "시간표 선택삭제") = vbNo Then
                    Set DBCmd = Nothing
                    Set DBRec = Nothing
                    
                    Exit Sub
                End If
            End If
        End If
            
            
        
        '## 계속 삭제진행
        Set DBRec = Nothing
        
        basDataBase.DBConn.BeginTrans
        
            
        sStr = ""
        sStr = sStr & "  DELETE"
        sStr = sStr & "    FROM SDTRX50TB"
        sStr = sStr & "   WHERE ACID   = '" & sAcID & "'"
        sStr = sStr & "     AND LSNCD  = '" & sLsnCD & "'"
        sStr = sStr & "     AND LESSON = " & sLesson
        sStr = sStr & "     AND WEEKS  = " & sWeeks
        
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
        
        DBCmd.Execute nExe, , -1
                
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nExe = 1 Then
            basDataBase.DBConn.CommitTrans
            
            
            '>> 다시 보여주기 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            '       sprTotSisu
            '       sprLsnSisu
            '       sprWork
            '       sprTimeTable
                cmdTotSisu.Tag = "REVIEW"
                    sprTotSisu.MaxRows = 0
                    Call Find_Sisu_TotalData        ' 계열시수내역 조회
                    
                    With sprTotSisu
                        .Row = fpsprTotSisu_Row.Value:  .Row2 = .Row
                        .Col = 1:                       .Col2 = 4
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Row = fpsprTotSisu_Row.Value:  .Row2 = .Row
                        .Col = 6:                       .Col2 = .MaxCols
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Col = .MaxCols:        .Value = 1
                    End With
                cmdTotSisu.Tag = ""
                Call cmdFindLsnSisu_Click       ' 해당 시수코드의 내용 조회
                Call cmdFindWork_Click          ' sprLsnSisu 부분의 선택 및 sprWork 에 대한 부분 조회
                Call cmdShowTimeTable_Click     ' sprTimeTable (전체 시간표 조회)
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간표 선택삭제"
            
        Else
            basDataBase.DBConn.RollbackTrans
            MsgBox "삭제 오류가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 선택삭제"
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    On Error Resume Next
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "선택 삭제시 에러가 발생하였습니다." & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "시간표 선택삭제"
    
    On Error GoTo 0
End Sub

























