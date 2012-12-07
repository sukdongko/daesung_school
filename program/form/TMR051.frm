VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form TMR051 
   Caption         =   "시간표 만들기 >> 전체시간표 구성"
   ClientHeight    =   13740
   ClientLeft      =   1830
   ClientTop       =   1305
   ClientWidth     =   19020
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   13740
   ScaleWidth      =   19020
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   5250
      TabIndex        =   55
      Top             =   30
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin FPSpread.vaSpread sprExcel 
      Height          =   9075
      Left            =   4230
      TabIndex        =   41
      Top             =   14190
      Visible         =   0   'False
      Width           =   14685
      _Version        =   393216
      _ExtentX        =   25903
      _ExtentY        =   16007
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
      SpreadDesigner  =   "TMR051.frx":0000
   End
   Begin VB.Frame fraAuto 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  '없음
      Height          =   13665
      Left            =   -3600
      TabIndex        =   37
      Top             =   4740
      Width           =   12285
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  '없음
         Height          =   13605
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   12225
         Begin FPSpread.vaSpread sprWork 
            Height          =   2835
            Left            =   7680
            TabIndex        =   18
            Top             =   4770
            Width           =   4005
            _Version        =   393216
            _ExtentX        =   7064
            _ExtentY        =   5001
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
            SpreadDesigner  =   "TMR051.frx":0218
         End
         Begin VB.CommandButton cmdWorkTamgu 
            Caption         =   "탐구 자유배정하기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7680
            TabIndex        =   57
            Top             =   4170
            Width           =   2895
         End
         Begin VB.ComboBox cboAutoTmrGbn 
            Height          =   300
            Left            =   240
            Style           =   2  '드롭다운 목록
            TabIndex        =   56
            Top             =   210
            Width           =   1965
         End
         Begin FPSpread.vaSpread sprAutoGwamokSort 
            Height          =   2595
            Left            =   7710
            TabIndex        =   53
            Top             =   8670
            Width           =   3705
            _Version        =   393216
            _ExtentX        =   6535
            _ExtentY        =   4577
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
            MaxCols         =   2
            MaxRows         =   5
            ProcessTab      =   -1  'True
            ScrollBars      =   0
            SpreadDesigner  =   "TMR051.frx":070C
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00F7EFE7&
            Caption         =   "등록"
            Height          =   225
            Left            =   5940
            TabIndex        =   16
            Top             =   690
            Width           =   675
         End
         Begin FPSpread.vaSpread sprAutoTeacher 
            Height          =   12885
            Left            =   0
            TabIndex        =   15
            Top             =   660
            Width           =   6825
            _Version        =   393216
            _ExtentX        =   12039
            _ExtentY        =   22728
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
            SpreadDesigner  =   "TMR051.frx":0CCA
         End
         Begin VB.CommandButton cmdWork 
            Caption         =   "자동 시간표 자동배분하기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7680
            TabIndex        =   17
            Top             =   3630
            Width           =   2895
         End
         Begin VB.CommandButton cmdCalcu_TCR 
            Caption         =   "강사현황"
            Height          =   405
            Left            =   2490
            TabIndex        =   14
            Top             =   150
            Width           =   1305
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   3930
            TabIndex        =   58
            Top             =   210
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   $"TMR051.frx":2789
            Height          =   2535
            Left            =   7080
            TabIndex        =   54
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label lblAutoClose 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫 기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   11010
            TabIndex        =   39
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin EditLib.fpMask fpYM 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
      _ExtentY        =   503
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
      BackColor       =   16777215
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
      Mask            =   "######"
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
   Begin VB.Frame frain 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Height          =   915
      Left            =   11130
      TabIndex        =   43
      Top             =   8580
      Width           =   3705
      Begin VB.Frame Frame6 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Height          =   855
         Left            =   30
         TabIndex        =   44
         Top             =   30
         Width           =   3645
         Begin VB.CommandButton cmdIpruck 
            Caption         =   "등 록"
            Height          =   315
            Left            =   2340
            TabIndex        =   59
            ToolTipText     =   "[과목,강사]  [과목,강사/계열/반명]"
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox txtinSpr 
            Enabled         =   0   'False
            Height          =   300
            Left            =   450
            TabIndex        =   48
            Text            =   "txtinSpr"
            Top             =   30
            Width           =   1815
         End
         Begin VB.TextBox txtData 
            Height          =   300
            Left            =   450
            TabIndex        =   47
            Text            =   "txtData"
            ToolTipText     =   "[과목,강사]  [과목,강사/계열/반명]"
            Top             =   390
            Width           =   1815
         End
         Begin VB.TextBox txtinCol 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2940
            TabIndex        =   46
            Text            =   "txtinCol"
            Top             =   30
            Width           =   555
         End
         Begin VB.TextBox txtinRow 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2340
            TabIndex        =   45
            Text            =   "txtinRow"
            Top             =   30
            Width           =   555
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "[과목,강사]    [과목,강사/계열/반명]"
            Height          =   210
            Left            =   450
            TabIndex        =   70
            ToolTipText     =   "[과목,강사]    [과목,강사/계열/반명]"
            Top             =   690
            Width           =   3165
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "입력"
            Height          =   210
            Left            =   60
            TabIndex        =   49
            Top             =   450
            Width           =   1185
         End
      End
   End
   Begin VB.Frame fraResult 
      BackColor       =   &H00000080&
      BorderStyle     =   0  '없음
      Height          =   3945
      Left            =   7470
      TabIndex        =   32
      Top             =   13530
      Width           =   9495
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   3825
         Left            =   60
         TabIndex        =   33
         Top             =   60
         Width           =   9375
         Begin RichTextLib.RichTextBox txtResult_LSN 
            Height          =   3255
            Left            =   60
            TabIndex        =   13
            Top             =   510
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5741
            _Version        =   393217
            TextRTF         =   $"TMR051.frx":2915
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "오류내역 보기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   270
            TabIndex        =   35
            Top             =   150
            Width           =   3135
         End
         Begin VB.Label lblClose 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫 기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   8160
            TabIndex        =   34
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdsprTmr_Tcr 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   20
      Top             =   9630
      Width           =   315
   End
   Begin VB.CommandButton cmdsprTmr_Lsn 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   19
      Top             =   510
      Width           =   315
   End
   Begin VB.CommandButton cmdSave_Tcr 
      Caption         =   "저장(강사)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3870
      TabIndex        =   6
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdViewResult 
      Caption         =   "오류보기"
      Height          =   315
      Left            =   6930
      TabIndex        =   7
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdSave_LSN 
      Caption         =   "저장(반)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2310
      TabIndex        =   5
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "조회"
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   30
      Width           =   945
   End
   Begin FPSpread.vaSpread sprTmr_Tcr 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   9570
      Width           =   14835
      _Version        =   393216
      _ExtentX        =   26167
      _ExtentY        =   7011
      _StockProps     =   64
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "TMR051.frx":29A5
   End
   Begin FPSpread.vaSpread sprTmr_Lsn 
      Height          =   9105
      Left            =   0
      TabIndex        =   21
      Top             =   480
      Width           =   11025
      _Version        =   393216
      _ExtentX        =   19447
      _ExtentY        =   16060
      _StockProps     =   64
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "TMR051.frx":6E3E
   End
   Begin VB.CommandButton cmdReCreatHeader 
      Caption         =   "헤더생성"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10080
      TabIndex        =   4
      Top             =   30
      Width           =   855
   End
   Begin VB.TextBox txtWeeks 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   10  '한글 
      Left            =   9420
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "txtWeeks"
      Top             =   45
      Width           =   525
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Height          =   13515
      Left            =   14850
      TabIndex        =   28
      Top             =   30
      Width           =   4155
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame4"
         Height          =   13455
         Left            =   60
         TabIndex        =   29
         Top             =   30
         Width           =   4065
         Begin VB.CommandButton cmdExcelToGwamok 
            Caption         =   "시간표 과목별로 엑셀자료"
            Height          =   435
            Left            =   60
            TabIndex        =   69
            Top             =   9090
            Width           =   2295
         End
         Begin VB.CommandButton cmdViewNotTeach 
            Caption         =   $"TMR051.frx":B2D7
            Height          =   495
            Left            =   60
            TabIndex        =   51
            Top             =   8040
            Width           =   1995
         End
         Begin VB.CommandButton cmdDelKME 
            Caption         =   $"TMR051.frx":B2F5
            Height          =   495
            Left            =   60
            TabIndex        =   50
            Top             =   7500
            Width           =   2295
         End
         Begin VB.CommandButton cmdTmrAllDelete 
            Caption         =   $"TMR051.frx":B317
            Height          =   495
            Left            =   60
            TabIndex        =   42
            Top             =   6930
            Width           =   2295
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "시간표 엑셀자료 만들기"
            Height          =   435
            Left            =   60
            TabIndex        =   40
            Top             =   8610
            Width           =   2295
         End
         Begin MSComDlg.CommonDialog dlgExcel 
            Left            =   2100
            Top             =   8880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdTmrChg 
            Caption         =   $"TMR051.frx":B335
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   60
            TabIndex        =   10
            Top             =   6360
            Width           =   2295
         End
         Begin VB.CommandButton cmdAutoTmr 
            Caption         =   $"TMR051.frx":B34E
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   60
            TabIndex        =   9
            Top             =   5790
            Width           =   2295
         End
         Begin FPSpread.vaSpread sprSubj 
            Height          =   7785
            Left            =   2490
            TabIndex        =   11
            Top             =   5610
            Width           =   1515
            _Version        =   393216
            _ExtentX        =   2672
            _ExtentY        =   13732
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
            MaxCols         =   2
            SpreadDesigner  =   "TMR051.frx":B36B
         End
         Begin FPSpread.vaSpread sprSisu 
            Height          =   2835
            Left            =   0
            TabIndex        =   24
            Top             =   2760
            Width           =   4005
            _Version        =   393216
            _ExtentX        =   7064
            _ExtentY        =   5001
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
            SpreadDesigner  =   "TMR051.frx":CBB1
         End
         Begin FPSpread.vaSpread sprGwamok 
            Height          =   2745
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   4005
            _Version        =   393216
            _ExtentX        =   7064
            _ExtentY        =   4842
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
            ScrollBars      =   2
            SpreadDesigner  =   "TMR051.frx":D0A5
         End
         Begin VB.Label lblNotTeaching 
            BackColor       =   &H00FF00FF&
            Height          =   285
            Left            =   2130
            TabIndex        =   52
            Top             =   8070
            Width           =   225
         End
         Begin VB.Label Label36 
            BackStyle       =   0  '투명
            Caption         =   "전체과목 내역"
            Height          =   210
            Left            =   1320
            TabIndex        =   30
            Top             =   5610
            Width           =   1185
         End
      End
   End
   Begin EditLib.fpLongInteger fpLesson 
      Height          =   285
      Left            =   8400
      TabIndex        =   0
      Top             =   45
      Width           =   525
      _Version        =   196608
      _ExtentX        =   926
      _ExtentY        =   503
      Enabled         =   0   'False
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
      MaxValue        =   "10"
      MinValue        =   "7"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   9525
      Left            =   11070
      TabIndex        =   25
      Top             =   30
      Width           =   3825
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   9465
         Left            =   30
         TabIndex        =   26
         Top             =   30
         Width           =   3765
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   60
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin VB.CommandButton cmdSearchTcr 
            Caption         =   "강사현황 조회"
            Height          =   315
            Left            =   0
            TabIndex        =   8
            Top             =   30
            Width           =   1755
         End
         Begin FPSpread.vaSpread sprTcr 
            Height          =   7635
            Left            =   0
            TabIndex        =   22
            Top             =   360
            Width           =   3765
            _Version        =   393216
            _ExtentX        =   6641
            _ExtentY        =   13467
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
            MaxCols         =   12
            SpreadDesigner  =   "TMR051.frx":EAB4
         End
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   1
            Left            =   480
            TabIndex        =   61
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   2
            Left            =   960
            TabIndex        =   62
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   3
            Left            =   1440
            TabIndex        =   63
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   4
            Left            =   1920
            TabIndex        =   64
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   5
            Left            =   2400
            TabIndex        =   65
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   6
            Left            =   2880
            TabIndex        =   66
            Top             =   8190
            Width           =   495
            _Version        =   196608
            _ExtentX        =   873
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
         Begin EditLib.fpDoubleSingle fpT 
            Height          =   345
            Index           =   7
            Left            =   3360
            TabIndex        =   67
            Top             =   8190
            Width           =   465
            _Version        =   196608
            _ExtentX        =   820
            _ExtentY        =   609
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
            DecimalPlaces   =   -1
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
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
            BackStyle       =   0  '투명
            Caption         =   "전     남     월     화     수     목     금     토"
            Height          =   210
            Left            =   180
            TabIndex        =   68
            Top             =   8010
            Width           =   3585
         End
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "[ 과목,강사 ] 넣어세요. 등록방법 -> 1,2 계열제외는 X01(3), 계열(1), 표시반명(10) 넣으세요. 입력후 반드시 엔터키를 치십시요."
      Height          =   180
      Left            =   0
      TabIndex        =   36
      Top             =   330
      Width           =   17295
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "요일"
      Height          =   210
      Left            =   8970
      TabIndex        =   31
      Top             =   105
      Width           =   465
   End
   Begin VB.Label Label23 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "교시"
      Height          =   210
      Left            =   7920
      TabIndex        =   27
      Top             =   105
      Width           =   465
   End
End
Attribute VB_Name = "TMR051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM026
'   모 듈  목 적 : 이동수업 시간표 등록
'
'   작   성   일 : 2008/01/04
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Const RowHeight = 12

Private Type tTmr
    TCRCD       As String
    TCRNM       As String
    SUBJCD      As String
    SUBJNM      As String
    
    WEEKS       As String
    LESSON      As String
    KAEYOL      As String
    KAEYOLNM    As String
    
    LSNCD       As String
    LSNNM       As String
    LSNCDNM     As String
    
End Type
Private uTmr()      As tTmr
Private nLesson_Max As Long
Private nOpenForm   As Integer

Private Type tTcr_Dup_Row_and_Col
    Row     As Long
    Col     As Long
End Type
Private uTcr_Dup_Row_and_Col() As tTcr_Dup_Row_and_Col


'< 헤더내용 처리 >
Private Sub Form_Activate()
    If nOpenForm = 0 Then
        With sprTmr_Lsn
            .Row = SpreadHeader + 1:        .RowHidden = True
            .Row = SpreadHeader + 3:        .RowHidden = True
            .Col = SpreadHeader + 1:        .ColHidden = True

            .AddCellSpan SpreadHeader, SpreadHeader, 3, 4

        End With

        If sprTmr_Tcr.RowHeaderCols < 2 Then Exit Sub
        With sprTmr_Tcr
            .Row = SpreadHeader + 1:        .RowHidden = True
            .Col = SpreadHeader:            .ColHidden = True
            .Col = SpreadHeader + 1:        .ColHidden = True

            .AddCellSpan SpreadHeader, SpreadHeader, 5, 3

        End With
        
        Call cmdFind_Click              '< 조회
        
        nOpenForm = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload LSN001
    Unload TMR011
    Unload TMR052
    
End Sub

Private Sub Form_Load()
    Dim ni      As Integer
    
    Me.Move 0, 0, 19140, 14210
    Me.KeyPreview = True
    
    ReDim uTmr(0) As tTmr
    ReDim uTcr_Dup_Row_and_Col(0) As tTcr_Dup_Row_and_Col
    
    fpYM.Text = Format(Now, "YYYYMM")
    
    nOpenForm = 0
    
    Me.Tag = "LOAD"
        Call Get_Max_Week_and_Lesson
        
        With sprSubj
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
            
            .Col = SpreadHeader
            .ColWidth(.Col) = 2
        End With
        
        With sprTcr
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
            Call Disp_Teacher_Sisu          '<< 강사자료 조회
            Call Disp_Subj
            
        End With
        
        With sprGwamok
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
            .MaxRows = 0
        End With
        
        With sprSisu
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
        End With
        
    '## 시간표 내역
        With sprTmr_Lsn
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2

            .Tag = "0"
            
            Call cmdReCreatHeader_Click     '<< 헤더생성
            
            .Move 0, 480, 11055, 9105
            .ZOrder 0
            
            cmdsprTmr_Lsn.Tag = "S"
            cmdsprTmr_Lsn.Top = .Top
            cmdsprTmr_Lsn.Left = .Left
            cmdsprTmr_Lsn.ZOrder 0
            
        End With
        
        With sprTmr_Tcr
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2

            .Tag = "0"
            
            .Move 0, 9600, 17385, 3945
            .ZOrder 0
            
            cmdsprTmr_Tcr.Tag = "S"
            cmdsprTmr_Tcr.Top = .Top
            cmdsprTmr_Tcr.Left = .Left
            cmdsprTmr_Tcr.ZOrder 0
            
        End With
        
        nLesson_Max = fpLesson.Value        '< 강의시간 max
        
        
        fraResult.Left = 60
        fraResult.Top = 60
        fraResult.Visible = False
        fraResult.ZOrder 0
        
        txtResult_LSN.Text = ""
        
        
        '-------------------------------------------
        ' 강사 시간표 자동배분하기
        '-------------------------------------------
            With sprWork
                .ShadowColor = basModule.ShadowColor2
                .ShadowDark = basModule.ShadowDark2
                .ShadowText = basModule.ShadowText2
                .GridColor = basModule.GridColor2
                .GrayAreaBackColor = basModule.GrayAreaBackColor2
    
                .Tag = "0"
            End With
            
            With sprAutoTeacher
                .ShadowColor = basModule.ShadowColor2
                .ShadowDark = basModule.ShadowDark2
                .ShadowText = basModule.ShadowText2
                .GridColor = basModule.GridColor2
                .GrayAreaBackColor = basModule.GrayAreaBackColor2
    
                .Tag = "0"
                
                .MaxRows = 0
            End With
            
            fraAuto.Top = 450
            fraAuto.Left = 2610
            fraAuto.Visible = False
        '-------------------------------------------
        
        
        cmdSave_LSN.Enabled = True
        cmdSave_Tcr.Enabled = False
        
        chkAll.Value = 0
        
        sprExcel.ZOrder 0
        
        txtinSpr.Text = ""
        txtinRow.Text = ""
        txtinCol.Text = ""
        txtData.Text = ""
        
        With cboAutoTmrGbn
            .Clear
            
            .AddItem "탐구영역" & Space(30) & "TAM"
            .AddItem "언/수/외" & Space(30) & "KME"
            .AddItem "전체" & Space(30) & "ALL"
            
            .ListIndex = 2
        End With
        
        For ni = 0 To 7 Step 1
            fpT(ni).Value = 0
        Next ni
        
    Me.Tag = ""
    
End Sub

Private Sub cmdViewResult_Click()
    fraResult.Visible = True
    
End Sub

Private Sub lblAutoClose_Click()
    fraAuto.Visible = False
    
End Sub

Private Sub lblClose_Click()
    fraResult.Visible = False
    
End Sub

Private Sub cmdsprTmr_Lsn_Click()

    If cmdsprTmr_Lsn.Tag = "S" Then
        sprTmr_Lsn.Move 0, 480, 18975, 13200
        sprTmr_Lsn.ZOrder 0
        
        cmdsprTmr_Lsn.Top = sprTmr_Lsn.Top
        cmdsprTmr_Lsn.Left = sprTmr_Lsn.Left
        
        cmdsprTmr_Lsn.Tag = "L"
        cmdsprTmr_Lsn.ZOrder 0
    Else
        sprTmr_Lsn.Move 0, 480, 11055, 9105
        sprTmr_Lsn.ZOrder 0
        
        cmdsprTmr_Lsn.Top = sprTmr_Lsn.Top
        cmdsprTmr_Lsn.Left = sprTmr_Lsn.Left
        
        cmdsprTmr_Lsn.Tag = "S"
        cmdsprTmr_Lsn.ZOrder 0
    End If
    
End Sub

Private Sub cmdsprTmr_Tcr_Click()
    If cmdsprTmr_Tcr.Tag = "S" Then
        sprTmr_Tcr.Move 0, 480, 18975, 13200
        sprTmr_Tcr.ZOrder 0
        
        cmdsprTmr_Tcr.Top = sprTmr_Tcr.Top
        cmdsprTmr_Tcr.Left = sprTmr_Tcr.Left
        
        cmdsprTmr_Tcr.Tag = "L"
        cmdsprTmr_Tcr.ZOrder 0
    Else
        sprTmr_Tcr.Move 0, 9600, 17385, 3945
        sprTmr_Tcr.ZOrder 0
        
        cmdsprTmr_Tcr.Top = sprTmr_Tcr.Top
        cmdsprTmr_Tcr.Left = sprTmr_Tcr.Left
        
        cmdsprTmr_Tcr.Tag = "S"
        cmdsprTmr_Tcr.ZOrder 0
    End If
    
End Sub


Private Sub cmdSearchTcr_Click()
    Dim nRow        As Long
    Dim nCol        As Long
    
    Call Disp_Teacher_Sisu
    Call Disp_Subj
    
    ' 초기화
    sprGwamok.MaxRows = 0
    With sprSisu
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
    End With
End Sub


'## 교시 요일 MAX값
Private Sub Get_Max_Week_and_Lesson()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT CASE WHEN CHKS = 6 THEN"
    sStr = sStr & "               '금'"
    sStr = sStr & "           ELSE CASE WHEN CHKS = 7 THEN"
    sStr = sStr & "               '토'"
    sStr = sStr & "           ELSE CASE WHEN CHKS = 8 THEN"
    sStr = sStr & "               '일'"
    sStr = sStr & "           ELSE"
    sStr = sStr & "               '금'"
    sStr = sStr & "           END END END CHKS,"
    sStr = sStr & "           MXLESSON"
    sStr = sStr & "      FROM (SELECT MAX(CHKS) AS CHKS, MAX(MXLESSON) AS MXLESSON"
    sStr = sStr & "              FROM ("
    sStr = sStr & "                    SELECT 6 AS CHKS, 7 AS MXLESSON"
    sStr = sStr & "                      FROM DUAL"
    sStr = sStr & "                    UNION ALL"
    sStr = sStr & "                    SELECT CASE WHEN MNWEEK  = 1 THEN"
    sStr = sStr & "                               8"
    sStr = sStr & "                           ELSE CASE WHEN MXWEEK  = 7 THEN"
    sStr = sStr & "                               7"
    sStr = sStr & "                           ELSE CASE WHEN MXWEEK <= 6 THEN"
    sStr = sStr & "                               6"
    sStr = sStr & "                           END END END CHKS,"
    sStr = sStr & "                           MXLESSON"
    sStr = sStr & "                      FROM (SELECT MAX(A.WEEKS) AS MXWEEK,"
    sStr = sStr & "                                   MIN(A.WEEKS) AS MNWEEK,"
    sStr = sStr & "                                   MAX(A.LESSON) AS MXLESSON"
    sStr = sStr & "                              FROM SDTRX50TB A, "
    
    sStr = sStr & "                                   (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                                      FROM SDLSN01TB "
    sStr = sStr & "                                     WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                    UNION"
    sStr = sStr & "                                    SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                      FROM SDLSN02TB "
    sStr = sStr & "                                     WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                   ) B"

    sStr = sStr & "                             WHERE A.ACID = B.ACID  "
    sStr = sStr & "                               AND A.LSNCD= B.LSNCD "
    sStr = sStr & "                               AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                               AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                            )"
    sStr = sStr & "                    )"
    sStr = sStr & "            )"
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    fpLesson.Value = 7
    txtWeeks.Text = "금"
    
    If DBRec.RecordCount = 1 Then
        txtWeeks.Text = Trim(DBRec.Fields("CHKS"))
        fpLesson.Value = CLng(DBRec.Fields("MXLESSON"))
    End If
   
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    
    fpLesson.Value = 7
    txtWeeks.Text = "금"
    
End Sub



'## 헤더생성
Private Sub cmdReCreatHeader_Click()

'<< 요일별 반 내역 SPREAD >>
    Call Weeks_And_Lesson_Header_Tmr            '< 요일 및 교시
    Call Lsn_And_Kaeyol_Header_Tmr              '< 반 내역
    
'<< 강사별 요일 내역 SPREAD >>
    
    Call Teachers_And_Subj_Header               '< 강사 및 과목
    Call Teachers_Weeks_And_Lesson_Header       '< 요일 및 교시
    
End Sub


'## 데이터 조회
Private Sub cmdFind_Click()

    txtResult_LSN.Text = ""
    
    If sprTmr_Lsn.MaxRows < 1 And sprTmr_Tcr.MaxRows < 1 Then Exit Sub
    
    If Me.Tag <> "LOAD" Then
        Call cmdSearchTcr_Click         '<< 강사현황 조회
        Call cmdReCreatHeader_Click     '<< 헤더생성
    End If
    
'<< 데이터조회 전 TYPE 변수에 넣음... >>
    If Data_into_TypeValue = False Then
        MsgBox "나타낼 시간표 자료가 없습니다.", vbExclamation + vbOKOnly, "시간표 조회"
    Else
        ' 조회된 데이터가 USER TYPE 변수에 들어 있어야 함.
        If UBound(uTmr) > 0 Then
            txtResult_LSN.Text = ""
            Call Disp_Tmr_Lsn                       '< 등록내역 조회 : 반별
            Call Disp_Tmr_Teacher                   '< 등록내역 조회 : 강사별
        End If
        
        If txtResult_LSN.Text > "" Then
            fraResult.Visible = True
        End If
        
        Call cmdViewNotTeach_Click
        
    End If
End Sub


'<< 요일별 반 내역 SPREAD >>
'   등록내역 조회
Private Sub Disp_Tmr_Lsn()
    Dim nRec        As Long
    Dim nRow        As Long
    Dim nCol        As Long
        
    Dim nKaeyol_Chg As Long     '< 계열 바뀌는 위치
    
    Dim sComp       As String
    Dim sTmp        As String
    
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim sKaeyol     As String
    Dim sLsnCD      As String
    
    Dim nSel_Row    As Long
    Dim nSel_Col    As Long
    
    Dim bRet        As Boolean
    
    With sprTmr_Lsn
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow
    End With
    
    With sprTmr_Lsn
        nKaeyol_Chg = 0
        
        For nCol = 1 To .MaxCols Step 1
            .Row = SpreadHeader + 3
            .Col = nCol
                If nCol = 1 Then sComp = Trim(.Text)
                sTmp = Trim(.Text)
            
            If StrComp(sComp, sTmp, vbTextCompare) <> 0 Then
                nKaeyol_Chg = .Col - 1              '< 계열이 바뀌는 column을 저장
                Exit For
            End If
        Next nCol
    End With
    
    'nLesson_Max        '< 교시 max 저장하고 있음. : 전역번수
    'nKaeyol_Chg        '< 계열 바뀌기 전 열(column)
    
    For nRec = 1 To UBound(uTmr) Step 1
    
        sWeek = "":         sLesson = ""
        sLsnCD = "":        sKaeyol = ""
                
        bRet = False
        
        If uTmr(nRec).WEEKS > "" And uTmr(nRec).LESSON > "" Then
        
            With sprTmr_Lsn
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                        .Col = SpreadHeader + 1:        sWeek = Trim(.Text)
                        .Col = SpreadHeader + 2:        sLesson = Trim(.Text)
                    
                    If StrComp(sWeek, uTmr(nRec).WEEKS, vbTextCompare) = 0 And _
                       StrComp(sLesson, uTmr(nRec).LESSON, vbTextCompare) = 0 Then
                       
                       ' row 결정..
                        nSel_Row = .Row
                       
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol
                                .Row = SpreadHeader + 1:    sLsnCD = Trim(.Text)
                                .Row = SpreadHeader + 3:    sKaeyol = Trim(.Text)
                            
                            If StrComp(sLsnCD, uTmr(nRec).LSNCD, vbTextCompare) = 0 And _
                               StrComp(sKaeyol, uTmr(nRec).KAEYOL, vbTextCompare) = 0 Then
                                
                                nSel_Col = .Col
                                
                                '<< 모든 자료가 있는 경우 >>
                                .Row = nSel_Row
                                .Col = nSel_Col
                                
                                If Trim(.Text) <> "" Then
                                    sTmp = Trim(.Text)
                                    sTmp = sTmp & "/" & uTmr(nRec).SUBJNM & "," & uTmr(nRec).TCRNM
                                        Call basFunction.Set_SprType_Text(sprTmr_Lsn, "center", "left", 60, sTmp)
                                        
                                    .Row2 = .Row
                                    .Col2 = .Col
                                    .BlockMode = True
                                        .BackColor = basModule.SectionColor1
                                        .BackColorStyle = BackColorStyleUnderGrid
                                    .BlockMode = False
                                Else
                                    sTmp = uTmr(nRec).SUBJNM & "," & uTmr(nRec).TCRNM
                                        Call basFunction.Set_SprType_Text(sprTmr_Lsn, "center", "left", 60, sTmp)
                                End If
                                
                                bRet = True         '< 정상처리
                                    
                            End If
                        Next nCol
                    End If
                Next nRow
            End With
        End If
    
                    
        '>> 정상처리 되지 않은 경우
        If bRet = False Then
            
            With uTmr(nRec)
                sTmp = ""
                sTmp = sTmp & "강사 [" & .TCRNM & ":" & .TCRCD & "]" & ", "
                sTmp = sTmp & "과목 [" & .SUBJNM & ":" & .SUBJCD & "]" & ", "
                                
                Select Case .WEEKS
                    Case "2"
                        sTmp = "월"
                    Case "3"
                        sTmp = "화"
                    Case "4"
                        sTmp = "수"
                    Case "5"
                        sTmp = "목"
                    Case "6"
                        sTmp = "금"
                    Case "7"
                        sTmp = "토"
                    Case "1"
                        sTmp = "일"
                End Select
                sTmp = sTmp & "요일 [" & sTmp & "] :" & ", "
                sTmp = sTmp & "교시 [" & .LESSON & "] :" & ", "
                sTmp = sTmp & "계열 [" & .KAEYOLNM & "] :" & ", "
                sTmp = sTmp & "반 [" & .LSNNM & "]"
            End With
    
            txtResult_LSN.Text = txtResult_LSN.Text & vbCrLf & sTmp
            
        End If
    
    Next nRec
    
End Sub


'<< 강사 요일별 내역 SPREAD >>
'   등록내역 조회
Private Sub Disp_Tmr_Teacher()
    Dim nRec        As Long
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sComp       As String
    Dim sTmp        As String
    
    Dim sTeacher    As String
    Dim sSubjCD     As String
    
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim nSel_Row    As Long
    Dim nSel_Col    As Long
    
    Dim bRet        As Boolean
    
    
    With sprTmr_Tcr
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow
    End With
    
    'nLesson_Max        '< 교시 max 저장하고 있음. : 전역번수
    
    For nRec = 1 To UBound(uTmr) Step 1
    
        sWeek = "":         sLesson = ""
        sTeacher = "":      sSubjCD = ""
                
        bRet = False
        
        If uTmr(nRec).WEEKS > "" And uTmr(nRec).LESSON > "" Then
        
            With sprTmr_Tcr
            
                '>> 강사
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                        .Col = SpreadHeader:            sTeacher = Trim(.Text)
                        .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
                        
                    If StrComp(sTeacher, uTmr(nRec).TCRCD, vbTextCompare) = 0 And _
                       StrComp(sSubjCD, uTmr(nRec).SUBJCD, vbTextCompare) = 0 Then
                       
                        nSel_Row = .Row
                        
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol
                                .Row = SpreadHeader + 1:    sWeek = Trim(.Text)
                                .Row = SpreadHeader + 2:    sLesson = Trim(.Text)
                                
                                If StrComp(sWeek, uTmr(nRec).WEEKS, vbTextCompare) = 0 And _
                                   StrComp(sLesson, uTmr(nRec).LESSON, vbTextCompare) = 0 Then
                                   
                                    nSel_Col = .Col
                                    
                                    '<< 모든 자료가 있는 경우 >>
                                    .Row = nSel_Row
                                    .Col = nSel_Col
                                    
                                    If Trim(.Text) > " " Then       '<< 중복자료 있는경우
                                        sTmp = Trim(.Text)
                                        sTmp = sTmp & "/" & uTmr(nRec).LSNCDNM
                                            Call basFunction.Set_SprType_Text(sprTmr_Tcr, "TOP", "left", 60, sTmp)
                                            
                                        .Row2 = .Row
                                        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.SectionColor1
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    Else
                                        sTmp = uTmr(nRec).LSNCDNM
                                            Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sTmp)
                                        
                                        If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                                           sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                                            ' no action
                                        Else
                                            .Row2 = .Row
                                            .Col2 = .Col
                                            .BlockMode = True
                                                .BackColor = basModule.WhiteColor
                                                .BackColorStyle = BackColorStyleUnderGrid
                                            .BlockMode = False
                                        End If
                                    End If
                                    
                                    
                                    bRet = True
                                    
                                End If
                            
                        Next nCol
                       
                    End If
                Next nRow
            End With
        End If
    
                    
        '>> 정상처리 되지 않은 경우
        If bRet = False Then
            
            With uTmr(nRec)
                sTmp = ""
                sTmp = sTmp & "강사 [" & .TCRNM & ":" & .TCRCD & "]" & ", "
                sTmp = sTmp & "과목 [" & .SUBJNM & ":" & .SUBJCD & "]" & ", "
                                
                Select Case .WEEKS
                    Case "2"
                        sTmp = "월"
                    Case "3"
                        sTmp = "화"
                    Case "4"
                        sTmp = "수"
                    Case "5"
                        sTmp = "목"
                    Case "6"
                        sTmp = "금"
                    Case "7"
                        sTmp = "토"
                    Case "1"
                        sTmp = "일"
                End Select
                sTmp = sTmp & "요일 [" & sTmp & "] :" & ", "
                sTmp = sTmp & "교시 [" & .LESSON & "] :" & ", "
                sTmp = sTmp & "계열 [" & .KAEYOLNM & "] :" & ", "
                sTmp = sTmp & "반 [" & .LSNNM & "]"
            End With
    
            txtResult_LSN.Text = txtResult_LSN.Text & vbCrLf & sTmp
            
        End If
    
    Next nRec
    
End Sub
    



'## 데이터조회 전 TYPE 변수에 넣음...
Private Function Data_into_TypeValue() As Boolean

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    
    Dim ni          As Long
    Dim nRec        As Long
    
    Dim bRet        As Boolean
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    
    On Error GoTo ErrStmt
    
    bRet = False
    ReDim uTmr(0) As tTmr       '< 초기화
    
    sStr = ""
    sStr = sStr & "        SELECT LSNCD     , NVL(LSNNM,'기타') AS LSNNM ,"
    sStr = sStr & "               KAEYOL    , KAEYOLNM    , CLASSNM, DAMIM, IDX, "
    sStr = sStr & "               NVL(LSNCDNM,'XXX') AS LSNCDNM,"
    sStr = sStr & "               TCRCD     , TCRNM     ,"
    sStr = sStr & "               SUBJCD    , SUBJNM    ,"
    sStr = sStr & "               WEEKS, LESSON"
    sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    
    Select Case Trim(basModule.SchCD)
        Case "N", "J"
            sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
        Case "S"
            sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
        Case "K"
            sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
    End Select
    
    sStr = sStr & "                       A.TCRCD, A.TCRNM,"
    sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
    sStr = sStr & "                       A.WEEKS, A.LESSON"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) A,"
    
    sStr = sStr & "                       (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                          FROM SDLSN01TB "
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'    sStr = sStr & "                        UNION"
'    sStr = sStr & "                        SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
'    sStr = sStr & "                          FROM SDLSN02TB "
'    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       ) B"

    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    
    Select Case Trim(basModule.SchCD)
        Case "N", "J"
            sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
        Case "S"
            sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
        Case "K"
            sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
    End Select
    
    sStr = sStr & "                       A.TCRCD, A.TCRNM ,"
    sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
    sStr = sStr & "                       A.WEEKS, A.LESSON"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN02TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT '00000' AS LSNCD, PRT_LSNNM AS LSNNM,"
    sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
    sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','인문계','2','자연계','기타') AS KAEYOLNM,"
    sStr = sStr & "                       '' AS CLASSNM,"
    sStr = sStr & "                       '' AS DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       PRT_LSN AS LSNCDNM,"
    sStr = sStr & "                       B.TCRCD, B.TCRNM,"
    sStr = sStr & "                       B.SUBJCD, B.SUBJNM,"
    sStr = sStr & "                       A.WEEKS, A.LESSON"
    sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                   AND A.LSNCD  = '00000'"
    sStr = sStr & "               )"
    sStr = sStr & "          ORDER BY TCRCD, SUBJCD "

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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    
    If DBRec.RecordCount = 0 Then
        
        '== 데이터 초기화 =====================================
        For nRow = 1 To sprTmr_Lsn.MaxRows Step 1
            For nCol = 1 To sprTmr_Lsn.MaxCols Step 1
                sprTmr_Lsn.Row = nRow
                sprTmr_Lsn.Col = nCol
                    sprTmr_Lsn.Text = ""
            Next nCol
        Next nRow
        
        For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
            For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
                sprTmr_Tcr.Row = nRow
                sprTmr_Tcr.Col = nCol
                    sprTmr_Tcr.Text = ""
            Next nCol
        Next nRow
        '======================================================
        
    Else
        
        DBRec.MoveFirst
        ReDim uTmr(DBRec.RecordCount) As tTmr           '<< 조회자료 type 변수
        
        For nRec = 1 To DBRec.RecordCount Step 1
    
            sTmp = "":      If IsNull(DBRec.Fields("TCRCD")) = False Then sTmp = DBRec.Fields("TCRCD")
                uTmr(nRec).TCRCD = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("TCRNM")) = False Then sTmp = DBRec.Fields("TCRNM")
                uTmr(nRec).TCRNM = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("SUBJCD")) = False Then sTmp = DBRec.Fields("SUBJCD")
                uTmr(nRec).SUBJCD = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("SUBJNM")) = False Then sTmp = DBRec.Fields("SUBJNM")
                uTmr(nRec).SUBJNM = sTmp
            
            sTmp = "":      If IsNull(DBRec.Fields("WEEKS")) = False Then sTmp = DBRec.Fields("WEEKS")
                uTmr(nRec).WEEKS = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("LESSON")) = False Then sTmp = DBRec.Fields("LESSON")
                uTmr(nRec).LESSON = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("KAEYOL")) = False Then sTmp = DBRec.Fields("KAEYOL")
                uTmr(nRec).KAEYOL = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("KAEYOLNM")) = False Then sTmp = DBRec.Fields("KAEYOLNM")
                uTmr(nRec).KAEYOLNM = sTmp
            
            sTmp = "":      If IsNull(DBRec.Fields("LSNCD")) = False Then sTmp = DBRec.Fields("LSNCD")
                uTmr(nRec).LSNCD = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("LSNNM")) = False Then sTmp = DBRec.Fields("LSNNM")
                uTmr(nRec).LSNNM = sTmp
            sTmp = "":      If IsNull(DBRec.Fields("LSNCDNM")) = False Then sTmp = DBRec.Fields("LSNCDNM")
                uTmr(nRec).LSNCDNM = sTmp
                        
            DBRec.MoveNext
        Next nRec
    End If
    
    bRet = True
    Data_into_TypeValue = bRet
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
   
    Data_into_TypeValue = bRet

End Function




'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 좌측 << 강사별 요일 내역 spread 조회 ################################################################################

Private Sub Teachers_Weeks_And_Lesson_Header()
    Dim nC_Header   As Integer
    Dim nL_Header   As Integer
    
    Dim nCol        As Long
    
    Dim nLesson     As Long
    Dim nWeeks      As Integer
    
    Dim sWeek       As String
    Dim sWeekCD     As String
    
    Select Case Trim(txtWeeks.Text)
        Case "월", "화", "수", "목", "금"
            nC_Header = 5
        Case "토"
            nC_Header = 6
        Case "일"
            nC_Header = 7
        Case Else
            nC_Header = 5
    End Select

    Select Case fpLesson.Value
        Case 10
            nL_Header = 11
        Case 9
            nL_Header = 10
        Case 8
            nL_Header = 9
        Case Is <= 7
            nL_Header = 8
    End Select

    With sprTmr_Tcr
        
        .MaxCols = nC_Header * nL_Header
        nWeeks = 1
        
        For nCol = 1 To .MaxCols Step nL_Header
            nWeeks = nWeeks + 1
            
            For nLesson = 1 To nL_Header Step 1
                Select Case nWeeks
                    Case 2
                        sWeekCD = "2":      sWeek = "월"
                    Case 3
                        sWeekCD = "3":      sWeek = "화"
                    Case 4
                        sWeekCD = "4":      sWeek = "수"
                    Case 5
                        sWeekCD = "5":      sWeek = "목"
                    Case 6
                        sWeekCD = "6":      sWeek = "금"
                    Case 7
                        sWeekCD = "7":      sWeek = "토"
                    Case 8
                        sWeekCD = "1":      sWeek = "일"
                End Select
                
                If sprTmr_Tcr.RowHeaderCols > 2 Then
                    .Col = nCol + nLesson - 1:  .ColWidth(.Col) = 2.7
                        .Row = SpreadHeader:        .Text = sWeek
                        .Row = SpreadHeader + 1:    .Text = sWeekCD
                        .Row = SpreadHeader + 2:    .Text = Trim(CStr(nLesson))
                End If
            Next nLesson
            
            .SetCellBorder .Col, 1, .Col, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
            
        Next nCol
        
        .Row = SpreadHeader
            .RowMerge = MergeAlways
            
    End With
    
End Sub

'<< 강사별 요일 내역 SPREAD >>
'   강사 및 과목
Private Sub Teachers_And_Subj_Header()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    Dim nRow        As Long
    
    On Error GoTo ErrStmt
    
    With sprTmr_Tcr
        If Me.Tag <> "LOAD" Then
            .Row = SpreadHeader:        .RowHidden = False
            .Row = SpreadHeader + 1:    .RowHidden = False
            .Row = SpreadHeader + 2:    .RowHidden = False
            
            .Col = SpreadHeader:        .ColHidden = False
            .Col = SpreadHeader + 1:    .ColHidden = False
            .Col = SpreadHeader + 2:    .ColHidden = False
            .Col = SpreadHeader + 3:    .ColHidden = False
            .Col = SpreadHeader + 4:    .ColHidden = False
            .Col = SpreadHeader + 5:    .ColHidden = False
            .Col = SpreadHeader + 6:    .ColHidden = False
        End If
        
        .MaxRows = 0
        .MaxCols = 0
        
        .ColHeaderRows = 1
        .RowHeaderCols = 1
    End With
    
    sStr = ""
    sStr = sStr & "        SELECT A.TCRCD , A.TCRNM ,"
    sStr = sStr & "               A.SUBJCD, "
    sStr = sStr & "               GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD) AS SUBJNM,"
    sStr = sStr & "               A.TCRGBN, "
    sStr = sStr & "               NVL(A.SISU,0) AS SISU, NVL(A.SISU,0)-NVL(B.SUM_SISU,0) AS SUM_SISU "
    sStr = sStr & "          FROM (SELECT A.ACID,"
    sStr = sStr & "                       A.TCRCD , MAX(A.TCRNM) AS TCRNM ,"
    sStr = sStr & "                       B.SUBJCD, "
    sStr = sStr & "                       SUM(B.SISU) AS SISU,"
    sStr = sStr & "                       DECODE(MAX(TCRGBN),'10','담임',"
    sStr = sStr & "                                          '20','강남출강',"
    sStr = sStr & "                                          '30','송파출강') AS TCRGBN"
    sStr = sStr & "                  FROM ("
    sStr = sStr & "                        SELECT ACID, TCRCD, MAX(SUBJCD) AS SUBJCD, MAX(TCRNM) AS TCRNM, MAX(TCRGBN) AS TCRGBN"
    sStr = sStr & "                          FROM SDTCR01TB"
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         GROUP BY ACID, TCRCD"
    sStr = sStr & "                        ) A, "
    sStr = sStr & "                       ("
    sStr = sStr & "                        SELECT A.ACID, A.TCRCD, A.SUBJCD, SUM(A.SISU) AS SISU"
    sStr = sStr & "                          FROM SDTCR11TB A, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) B"

    sStr = sStr & "                         WHERE A.ACID  = B.ACID "
    sStr = sStr & "                           AND A.LSNCD = B.LSNCD "
    
    sStr = sStr & "                           AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    
    sStr = sStr & "                         GROUP BY A.ACID, A.TCRCD, A.SUBJCD"
    sStr = sStr & "                        ) B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.TCRCD = B.TCRCD"
    sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                 GROUP BY A.ACID, A.TCRCD, B.SUBJCD"
    sStr = sStr & "                ) A,"
    sStr = sStr & "               (SELECT ACID, TCRCD, SUBJCD, SUM(SUN)+SUM(MON)+SUM(TUE)+SUM(WED)+SUM(THU)+SUM(FRI)+SUM(SAT) AS SUM_SISU"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.TCRCD, A.SUBJCD,"
    sStr = sStr & "                               DECODE(A.WEEKS, 1, 1, 0) AS SUN,            /* 일요일 */"
    sStr = sStr & "                               DECODE(A.WEEKS, 2, 1, 0) AS MON,"
    sStr = sStr & "                               DECODE(A.WEEKS, 3, 1, 0) AS TUE,"
    sStr = sStr & "                               DECODE(A.WEEKS, 4, 1, 0) AS WED,"
    sStr = sStr & "                               DECODE(A.WEEKS, 5, 1, 0) AS THU,"
    sStr = sStr & "                               DECODE(A.WEEKS, 6, 1, 0) AS FRI,"
    sStr = sStr & "                               DECODE(A.WEEKS, 7, 1, 0) AS SAT             /* 토요일 */"
    sStr = sStr & "                          FROM SDTRX50TB A, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) B"

    sStr = sStr & "                         WHERE A.ACID  = B.ACID  "
    sStr = sStr & "                           AND A.LSNCD = B.LSNCD "
    sStr = sStr & "                           AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        )"
    sStr = sStr & "                 GROUP BY ACID, TCRCD, SUBJCD"
    sStr = sStr & "                ) B"
    sStr = sStr & "         WHERE A.ACID = B.ACID (+)"
    sStr = sStr & "           AND A.TCRCD = B.TCRCD (+)"
    sStr = sStr & "           AND A.SUBJCD = B.SUBJCD (+)"
    sStr = sStr & "         ORDER BY TCRCD, SUBJCD"
    
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    If DBRec.RecordCount < 1 Then
        MsgBox "강사를 등록하여 주십시요.", vbExclamation + vbOKOnly, "강사 헤더생성"
        Load TMR011
        TMR011.Show
        TMR011.ZOrder 0
        
        Exit Sub
    ElseIf DBRec.RecordCount > 0 Then
        
        sprTmr_Tcr.ColHeaderRows = 3        '< 요일
        sprTmr_Tcr.RowHeaderCols = 7        '< 강사
        
        '<< 헤더생성 >> -----------------------------------------------------------------------
        sprTmr_Tcr.Row = SpreadHeader + 1:        sprTmr_Tcr.RowHidden = True
        'sprTmr_Tcr.Col = SpreadHeader:            sprTmr_Tcr.ColHidden = True
        'sprTmr_Tcr.Col = SpreadHeader + 1:        sprTmr_Tcr.ColHidden = True

        sprTmr_Tcr.AddCellSpan SpreadHeader, SpreadHeader, 5, 3
        '--------------------------------------------------------------------------------------

        DBRec.MoveFirst
        sprTmr_Tcr.MaxRows = DBRec.RecordCount
        
        For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
            sprTmr_Tcr.Row = nRow
                
                sprTmr_Tcr.Col = SpreadHeader:      sprTmr_Tcr.Text = Trim(DBRec.Fields("TCRCD")):      sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 4
                sprTmr_Tcr.Col = SpreadHeader + 1:
                    If IsNull(DBRec.Fields("SUBJCD")) = True Then
                        sprTmr_Tcr.Text = "":                               sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 3
                    Else
                        sprTmr_Tcr.Text = Trim(DBRec.Fields("SUBJCD")):     sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 3
                    End If
                sprTmr_Tcr.Col = SpreadHeader + 2:  sprTmr_Tcr.Text = Trim(DBRec.Fields("TCRNM")):      sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 6
                sprTmr_Tcr.Col = SpreadHeader + 3:  sprTmr_Tcr.Text = Trim(DBRec.Fields("SUBJNM")):     sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 5
                sprTmr_Tcr.Col = SpreadHeader + 4:  sprTmr_Tcr.Text = Trim(DBRec.Fields("SUBJNM")):     sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 4
                    If IsNull(DBRec.Fields("TCRGBN")) = True Then
                        sprTmr_Tcr.Text = " "
                    Else
                        sprTmr_Tcr.Text = Trim(DBRec.Fields("TCRGBN"))
                    End If
            
                sprTmr_Tcr.Col = SpreadHeader + 5
                    If Trim(DBRec.Fields("SISU")) = "0" Then
                        sprTmr_Tcr.Text = " "
                    Else
                        sprTmr_Tcr.Text = Trim(DBRec.Fields("SISU"))
                    End If
                    sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 3
                sprTmr_Tcr.Col = SpreadHeader + 6
                    If Trim(DBRec.Fields("SUM_SISU")) = "0" Then
                        sprTmr_Tcr.Text = " "
                    Else
                        sprTmr_Tcr.Text = Trim(DBRec.Fields("SUM_SISU"))
                    End If
                    sprTmr_Tcr.ColWidth(sprTmr_Tcr.Col) = 3
                
            DBRec.MoveNext
        Next nRow
    End If
   
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    
    MsgBox "강사를 등록하여 주십시요.", vbExclamation + vbOKOnly, "강사 헤더생성"
    
    Load TMR011
    TMR011.Show
    TMR011.ZOrder 0

End Sub


'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 좌측 << 요일별 반 내역 spread 조회 ##################################################################################

'<< 요일별 반 내역 SPREAD >>
'   반 내역
Private Sub Lsn_And_Kaeyol_Header_Tmr()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    Dim nCol        As Long
    
    Dim sKaeyol     As String
    Dim sWeek       As String
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL "
    sStr = sStr & "      FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL "
    sStr = sStr & "              FROM SDLSN01TB "
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    
'    sStr = sStr & "            UNION ALL "                                      '2009.01.12 추가
'    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL "
'    sStr = sStr & "              FROM SDLSN02TB "
'    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    
    sStr = sStr & "            UNION ALL"
    sStr = sStr & "            SELECT '" & Trim(basModule.SchCD) & "' AS ACID, '00000' AS LSNCD, '기타' AS LSNNM, 'ZZ' AS LSNCDNM, '01' AS KAEYOL"
    sStr = sStr & "              FROM DUAL"
    sStr = sStr & "            UNION ALL"
    sStr = sStr & "            SELECT '" & Trim(basModule.SchCD) & "' AS ACID, '00000' AS LSNCD, '기타' AS LSNNM, 'ZZ' AS LSNCDNM, '02' AS KAEYOL"
    sStr = sStr & "              FROM DUAL"
    sStr = sStr & "            UNION ALL"
    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL"
    sStr = sStr & "              FROM SDLSN02TB"
    sStr = sStr & "             WHERE (ACID, LSNCD)"
    sStr = sStr & "                IN (SELECT ACID, LSNCD"
    sStr = sStr & "                      FROM (SELECT ACID, LSNCD,"
    sStr = sStr & "                                   CASE WHEN (TAMGU1 +"
    sStr = sStr & "                                              TAMGU2 +"
    sStr = sStr & "                                              TAMGU3 +"
    sStr = sStr & "                                              TAMGU4 +"
    sStr = sStr & "                                              TAMGU5 +"
    sStr = sStr & "                                              TAMGU6 +"
    sStr = sStr & "                                              TAMGU7 +"
    sStr = sStr & "                                              TAMGU8 +"
    sStr = sStr & "                                              TAMGU9 +"
    sStr = sStr & "                                              TAMGU10+"
    sStr = sStr & "                                              TAMGU11+"
    sStr = sStr & "                                              J2SEL  +"
    sStr = sStr & "                                              NONSUL1+"
    sStr = sStr & "                                              NONSUL2+"
    sStr = sStr & "                                              NONSUL3+"
    sStr = sStr & "                                              NONSUL4) > 0 THEN"
    sStr = sStr & "                                       1"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       0"
    sStr = sStr & "                                   END INWON,"
    sStr = sStr & "                                   CASE WHEN (DECODE(TAMGU_CL1  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL2  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL3  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL4  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL5  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL6  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL7  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL8  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL9  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL10 , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL11 , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(J2SEL_CL   , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL1_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL2_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL3_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL4_CL , 16777215, 0, 1)) > 0 THEN"
    sStr = sStr & "                                       1"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       0"
    sStr = sStr & "                                   END NCOL"
    sStr = sStr & "                              FROM SDLSN05TB"
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                            )"
    sStr = sStr & "                     WHERE (INWON > 0 OR NCOL > 0)"
    sStr = sStr & "                       AND LSNCD >= '90000'"
    sStr = sStr & "                    )"
    sStr = sStr & "            )"
    sStr = sStr & "     ORDER BY KAEYOL, LSNCDNM"
    
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sKaeyol = ""
    
    If DBRec.RecordCount < 1 Then
        MsgBox "반을 등록하여 주십시요.", vbExclamation + vbOKOnly, "반 헤더생성"
        Load LSN001
        LSN001.Show
        LSN001.ZOrder 0
        
        Exit Sub
    ElseIf DBRec.RecordCount > 0 Then
        
        DBRec.MoveFirst
        sprTmr_Lsn.MaxCols = DBRec.RecordCount
        
        sKaeyol = ""
            
        For nCol = 1 To sprTmr_Lsn.MaxCols Step 1
            
            If nCol = 1 Then sKaeyol = Trim(DBRec.Fields("KAEYOL"))
            
            sprTmr_Lsn.Col = nCol
                
                sprTmr_Lsn.Row = SpreadHeader:      sprTmr_Lsn.Text = Trim(DBRec.Fields("LSNNM"))
                sprTmr_Lsn.Row = SpreadHeader + 1:  sprTmr_Lsn.Text = Trim(DBRec.Fields("LSNCD"))
                sprTmr_Lsn.Row = SpreadHeader + 2:  sprTmr_Lsn.Text = Trim(DBRec.Fields("LSNCDNM"))
                sprTmr_Lsn.Row = SpreadHeader + 3:  sprTmr_Lsn.Text = Trim(DBRec.Fields("KAEYOL"))
                
            If StrComp(sKaeyol, Trim(DBRec.Fields("KAEYOL")), vbTextCompare) <> 0 Then
                sprTmr_Lsn.SetCellBorder sprTmr_Lsn.Col, 1, sprTmr_Lsn.Col, sprTmr_Lsn.MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
                sKaeyol = Trim(DBRec.Fields("KAEYOL"))
            End If
            
            DBRec.MoveNext
        Next nCol
    End If
    
    sWeek = ""
    If sprTmr_Lsn.MaxRows > 0 Then
        For ni = 1 To sprTmr_Lsn.MaxRows Step 1
        
            sprTmr_Lsn.Row = ni
            sprTmr_Lsn.Col = SpreadHeader + 1
            If ni = 1 Then sWeek = Trim(sprTmr_Lsn.Text)
            
            If StrComp(sWeek, Trim(sprTmr_Lsn.Text), vbTextCompare) <> 0 Then
                sprTmr_Lsn.SetCellBorder 1, sprTmr_Lsn.Row, sprTmr_Lsn.MaxCols, sprTmr_Lsn.Row, 4, basModule.SectionColor1, CellBorderStyleSolid
                sWeek = Trim(sprTmr_Lsn.Text)
            End If
            
        Next ni
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    
    MsgBox "반을 등록하여 주십시요.", vbExclamation + vbOKOnly, "반 헤더생성"
    
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub

'<< 요일별 반 내역 SPREAD >>
'   요일 및 교시                <<= Get_Max_Week_and_Lesson 참조하세요...
Private Sub Weeks_And_Lesson_Header_Tmr()
    Dim nR_Header   As Integer
    Dim nL_Header   As Integer
    
    Dim nRow        As Long
    'Dim nCol        As Long
    
    Dim nLesson     As Long
    Dim nWeeks      As Integer
    
    Dim sWeek       As String
    Dim sWeekCD     As String
    
    With sprTmr_Lsn
        If Me.Tag <> "LOAD" Then
            .Row = SpreadHeader:        .RowHidden = False
            .Row = SpreadHeader + 1:    .RowHidden = False
            .Row = SpreadHeader + 2:    .RowHidden = False
            .Row = SpreadHeader + 3:    .RowHidden = False
            
            .Col = SpreadHeader:        .ColHidden = False
            .Col = SpreadHeader + 1:    .ColHidden = False
            .Col = SpreadHeader + 2:    .ColHidden = False
        End If
        
        .MaxRows = 0
        .MaxCols = 0
        
        .ColHeaderRows = 1
        .RowHeaderCols = 1
        
    End With
    
    Select Case Trim(txtWeeks.Text)
        Case "월", "화", "수", "목", "금"
            nR_Header = 5
        Case "토"
            nR_Header = 6
        Case "일"
            nR_Header = 7
        Case Else
            nR_Header = 5
    End Select

    Select Case fpLesson.Value
        Case 10
            nL_Header = 11
        Case 9
            nL_Header = 10
        Case 8
            nL_Header = 9
        Case Is <= 7
            nL_Header = 8
    End Select

    With sprTmr_Lsn
        .ColHeaderRows = 4
        .RowHeaderCols = 3
        
        '<< 헤더처리 >> ----------------------------------------------------
        .Row = SpreadHeader + 1:        .RowHidden = True
        .Row = SpreadHeader + 3:        .RowHidden = True
        .Col = SpreadHeader + 1:        .ColHidden = True

        .AddCellSpan SpreadHeader, SpreadHeader, 3, 4
        '-------------------------------------------------------------------
        
        .MaxRows = nR_Header * nL_Header
        nWeeks = 1
        
        For nRow = 1 To .MaxRows Step nL_Header
            nWeeks = nWeeks + 1
            
            For nLesson = 1 To nL_Header Step 1
                Select Case nWeeks
                    Case 2
                        sWeekCD = "2":      sWeek = "월"
                    Case 3
                        sWeekCD = "3":      sWeek = "화"
                    Case 4
                        sWeekCD = "4":      sWeek = "수"
                    Case 5
                        sWeekCD = "5":      sWeek = "목"
                    Case 6
                        sWeekCD = "6":      sWeek = "금"
                    Case 7
                        sWeekCD = "7":      sWeek = "토"
                    Case 8
                        sWeekCD = "1":      sWeek = "일"
                End Select
                
                .Row = nRow + nLesson - 1
                    .Col = SpreadHeader:        .Text = sWeek
                    .Col = SpreadHeader + 1:    .Text = sWeekCD
                    .Col = SpreadHeader + 2:    .Text = Trim(CStr(nLesson))
            Next nLesson
        Next nRow
        
        .Col = SpreadHeader
            .ColMerge = MergeAlways
    End With
    
End Sub







































'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------




'############################# 우측 << 1. 강사시수내역 / 2. 선택강사의 과목 및 영역 / 3. 선택강사 시간표 ##################################################################

'## 강사별 시수내역 조회
Private Sub Disp_Teacher_Sisu()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT A.ACID, A.TCRCD, A.TCRNM, "
    sStr = sStr & "           NVL(A.SISU,0) AS SISU,"
    sStr = sStr & "           NVL(B.SUM_SISU,0) AS SUM_SISU,"
    
    sStr = sStr & "           NVL(A.SISU,0)-NVL(B.SUM_SISU,0) AS CHA_SISU,"
    
    sStr = sStr & "           NVL(B.SUN,0) AS SUN,"
    sStr = sStr & "           NVL(B.MON,0) AS MON,"
    sStr = sStr & "           NVL(B.TUE,0) AS TUE,"
    sStr = sStr & "           NVL(B.WED,0) AS WED,"
    sStr = sStr & "           NVL(B.THU,0) AS THU,"
    sStr = sStr & "           NVL(B.FRI,0) AS FRI,"
    sStr = sStr & "           NVL(B.SAT,0) AS SAT"
    sStr = sStr & "      FROM ("
    sStr = sStr & "            SELECT A.ACID, A.TCRCD, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, SISU"
    sStr = sStr & "              FROM (SELECT ACID, TCRCD"
    sStr = sStr & "                      FROM SDTCR01TB"
    sStr = sStr & "                     WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                     GROUP BY ACID, TCRCD"
    sStr = sStr & "                    ) A,"
    sStr = sStr & "                   (SELECT A.ACID, A.TCRCD, SUM(A.SISU) AS SISU"
    sStr = sStr & "                      FROM SDTCR11TB A, "
    sStr = sStr & "                           (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "      '2009.01.12 추가
    sStr = sStr & "                              FROM SDLSN01TB "
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                            UNION"
    sStr = sStr & "                            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                              FROM SDLSN02TB "
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                           ) B"
    sStr = sStr & "                     WHERE A.ACID  = B.ACID "
    sStr = sStr & "                       AND A.LSNCD = B.LSNCD "
    sStr = sStr & "                       AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                     GROUP BY A.ACID, A.TCRCD"
    sStr = sStr & "                   ) B"
    sStr = sStr & "             WHERE A.ACID = B.ACID"
    sStr = sStr & "               AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "               AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    
    sStr = sStr & "            ) A,"
    sStr = sStr & "           (SELECT ACID, TCRCD,"
    sStr = sStr & "                   SUM(SUN)+SUM(MON)+SUM(TUE)+SUM(WED)+SUM(THU)+SUM(FRI)+SUM(SAT) AS SUM_SISU,"
    sStr = sStr & "                   SUM(SUN) AS SUN,"
    sStr = sStr & "                   SUM(MON) AS MON,"
    sStr = sStr & "                   SUM(TUE) AS TUE,"
    sStr = sStr & "                   SUM(WED) AS WED,"
    sStr = sStr & "                   SUM(THU) AS THU,"
    sStr = sStr & "                   SUM(FRI) AS FRI,"
    sStr = sStr & "                   SUM(SAT) AS SAT"
    sStr = sStr & "              FROM (SELECT A.ACID, A.TCRCD,"
    sStr = sStr & "                           DECODE(A.WEEKS, 1, 1, 0) AS SUN,          /* 일요일 */"
    sStr = sStr & "                           DECODE(A.WEEKS, 2, 1, 0) AS MON,"
    sStr = sStr & "                           DECODE(A.WEEKS, 3, 1, 0) AS TUE,"
    sStr = sStr & "                           DECODE(A.WEEKS, 4, 1, 0) AS WED,"
    sStr = sStr & "                           DECODE(A.WEEKS, 5, 1, 0) AS THU,"
    sStr = sStr & "                           DECODE(A.WEEKS, 6, 1, 0) AS FRI,"
    sStr = sStr & "                           DECODE(A.WEEKS, 7, 1, 0) AS SAT          /* 토요일 */"
    sStr = sStr & "                      FROM SDTRX50TB A, "
    sStr = sStr & "                           (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
    sStr = sStr & "                              FROM SDLSN01TB "
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                            UNION"
    sStr = sStr & "                            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                              FROM SDLSN02TB "
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    
    sStr = sStr & "                            Union"
    sStr = sStr & "                            SELECT '" & Trim(basModule.SchCD) & "' AS ACID, '00000' AS LSNCD, '' AS LSNNM, '' LSNCDNM, '' KAEYOL, '' DAMIM, '' BASE_CLASS"
    sStr = sStr & "                              From DUAL"
    
    sStr = sStr & "                           ) B"
    sStr = sStr & "                     WHERE A.ACID  = B.ACID "
    sStr = sStr & "                       AND A.LSNCD = B.LSNCD "
    sStr = sStr & "                       AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                       AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                    )"
    sStr = sStr & "             GROUP BY ACID, TCRCD"
    sStr = sStr & "            ) B"
    sStr = sStr & "     WHERE A.ACID  = B.ACID (+)"
    sStr = sStr & "       AND A.TCRCD = B.TCRCD (+)"
    sStr = sStr & "     ORDER BY TCRCD "
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    For ni = 0 To 7 Step 1
        fpT(ni).Value = 0
    Next ni
    
    sprTcr.MaxRows = 0
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows
                
                sprTcr.Col = 1
                    sTmp = " ":     If IsNull(.Fields("ACID")) = False Then sTmp = Trim(.Fields("ACID"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprTcr.Col = sprTcr.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprTcr.SetCellBorder sprTcr.Col, 1, sprTcr.Col, sprTcr.MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("SISU")) = True Then nTmp = CLng(.Fields("SISU"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(0).Value = fpT(0).Value + nTmp
                        
                sprTcr.SetCellBorder sprTcr.Col, 1, sprTcr.Col, sprTcr.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("CHA_SISU")) = True Then nTmp = CLng(.Fields("CHA_SISU"))
                        If nTmp <> 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", nTmp)
                        
                        If nTmp < 0 Then
                            sprTcr.ForeColor = basModule.SectionColor1
                        Else
                            sprTcr.ForeColor = basModule.SectionColor2
                        End If
                        
                    fpT(1).Value = fpT(1).Value + nTmp
                    
                sprTcr.SetCellBorder sprTcr.Col, 1, sprTcr.Col, sprTcr.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
                        
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("MON")) = True Then nTmp = CLng(.Fields("MON"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(2).Value = fpT(2).Value + nTmp
                    
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("TUE")) = True Then nTmp = CLng(.Fields("TUE"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(3).Value = fpT(3).Value + nTmp
                    
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("WED")) = True Then nTmp = CLng(.Fields("WED"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(4).Value = fpT(4).Value + nTmp
                
                sprTcr.SetCellBorder sprTcr.Col, 1, sprTcr.Col, sprTcr.MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("THU")) = True Then nTmp = CLng(.Fields("THU"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(5).Value = fpT(5).Value + nTmp
                    
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("FRI")) = True Then nTmp = CLng(.Fields("FRI"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(6).Value = fpT(6).Value + nTmp
                    
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("SAT")) = True Then nTmp = CLng(.Fields("SAT"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                    fpT(7).Value = fpT(7).Value + nTmp
                    
                sprTcr.Col = sprTcr.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("SUN")) = True Then nTmp = CLng(.Fields("SUN"))
                        If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTcr, 0, 0, 99999, "", nTmp)
                
                .MoveNext
            Next nRec
        End If
    End With
    
    With sprTcr
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "각 강사별 시수내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사별 시수내역 조회"
End Sub


'## 전체과목내역
Private Sub Disp_Subj()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT SUBJCD, SUBJNM"
    sStr = sStr & "      From SDTCR01TB"
    sStr = sStr & "     WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     GROUP BY SUBJCD, SUBJNM"
    sStr = sStr & "     ORDER BY SUBJCD"
        
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sprSubj.MaxRows = 0
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprSubj.MaxRows = sprSubj.MaxRows + 1
                sprSubj.Row = sprSubj.MaxRows
                
                sprSubj.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSubj.Col = sprSubj.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                .MoveNext
            Next nRec
        End If
    End With
    
    With sprSubj
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "전체과목 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체과목 조회"
End Sub




'## 강사 선택 : 강사 선택시 상세내역 조회
Private Sub sprTcr_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sAcID       As String
    Dim sTcrCD      As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprTcr
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):      .Row2 = .Row
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
        
        .Row = Row
        .Col = 1:       sAcID = Trim(.Text)
        .Col = 2:       sTcrCD = Trim(.Text)
        
       '## 선택 강사의 강의과목
        Call Disp_Tcr_Gwamok(sAcID, sTcrCD)
        
       '## 선택 강사의 시수내역 조회
        Call Disp_Tcr_Sisu(sAcID, sTcrCD)
        
    End With


End Sub

'## 선택 강사의 강의과목
Private Sub Disp_Tcr_Gwamok(ByVal aAcID As String, ByVal aTcrCD As String)

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT GET_TCRNM(ACID,TCRCD) AS TCRNM,"
    sStr = sStr & "        GET_SUBJNM(ACID,TCRCD,SUBJCD) AS SUBJNM,"
    sStr = sStr & "        GET_LSNNM(ACID,LSNCD) AS LSNNM,"
    sStr = sStr & "        GET_KEAYOL_N_LSN_TCR01(ACID,LSNCD) AS LSNCDNM,"
    sStr = sStr & "        TT, SISU, TSISU"
    sStr = sStr & "   FROM (SELECT ACID, TCRCD, '00' AS SUBJCD, '00000' AS LSNCD, '' AS LSNCDNM,"
    sStr = sStr & "                SUM(TSISU) AS TSISU, SUM(SISU) AS SISU,"
    sStr = sStr & "                SUM(TSISU)-SUM(SISU) AS TT"
    sStr = sStr & "           FROM (SELECT A.ACID, A.TCRCD, A.SUBJCD, B.LSNCD, NVL(B.SISU,0) AS TSISU, 0 AS SISU"
    sStr = sStr & "                   FROM SDTCR01TB A, SDTCR11TB B, "
    
    sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                           FROM SDLSN01TB "
    sStr = sStr & "                           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         UNION"
    sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                           FROM SDLSN02TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         ) C"
    
    sStr = sStr & "                  WHERE A.ACID   = B.ACID"
    sStr = sStr & "                    AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                    AND A.SUBJCD = B.SUBJCD"
    
    sStr = sStr & "                    AND B.ACID   = C.ACID  "
    sStr = sStr & "                    AND B.LSNCD  = C.LSNCD "
    
    sStr = sStr & "                    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                    AND A.TCRCD  = '" & aTcrCD & "'"
    sStr = sStr & "                 UNION ALL"
    sStr = sStr & "                 SELECT A.ACID, A.TCRCD, A.SUBJCD, A.LSNCD, 0 AS TSISU, SUM(NVL(A.SISU,0)) AS SISU"
    sStr = sStr & "                   FROM SDTRX50TB A, "
    
    sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                           FROM SDLSN01TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         UNION"
    sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                           FROM SDLSN02TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                           ) B "
    
    sStr = sStr & "                  WHERE A.ACID  = B.ACID  "
    sStr = sStr & "                    AND A.LSNCD = B.LSNCD "
    sStr = sStr & "                    AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                    AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                    AND A.TCRCD = '" & aTcrCD & "'"
    sStr = sStr & "                  GROUP BY A.YM, A.ACID, A.TCRCD, A.SUBJCD, A.LSNCD"
    sStr = sStr & "                 )"
    sStr = sStr & "          GROUP BY ACID, TCRCD"
    sStr = sStr & "         UNION ALL"
    sStr = sStr & "         SELECT ACID, TCRCD, SUBJCD, LSNCD, GET_LSNCDNM(ACID, LSNCD) AS LSNCDNM,"
    sStr = sStr & "                SUM(TSISU) AS TSISU, SUM(SISU) AS SISU,"
    sStr = sStr & "                SUM(TSISU)-SUM(SISU) AS TT"
    sStr = sStr & "           FROM (SELECT A.ACID, A.TCRCD, A.SUBJCD, B.LSNCD, NVL(B.SISU,0) AS TSISU, 0 AS SISU"
    sStr = sStr & "                   FROM SDTCR01TB A, SDTCR11TB B, "
    
    sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                           FROM SDLSN01TB "
    sStr = sStr & "                           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         UNION"
    sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                           FROM SDLSN02TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         ) C"

    sStr = sStr & "                  WHERE A.ACID   = B.ACID"
    sStr = sStr & "                    AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                    AND A.SUBJCD = B.SUBJCD"
    
    sStr = sStr & "                    AND B.ACID   = C.ACID  "
    sStr = sStr & "                    AND B.LSNCD  = C.LSNCD "
    
    sStr = sStr & "                    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                    AND A.TCRCD  = '" & aTcrCD & "'"
    sStr = sStr & "                 UNION ALL"
    sStr = sStr & "                 SELECT A.ACID, A.TCRCD, A.SUBJCD, A.LSNCD, 0 AS TSISU, SUM(NVL(A.SISU,0)) AS SISU"
    sStr = sStr & "                   FROM SDTRX50TB A, "
    
    sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                           FROM SDLSN01TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         UNION"
    sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                           FROM SDLSN02TB "
    sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) B"
    
    sStr = sStr & "                  WHERE A.ACID  = B.ACID  "
    sStr = sStr & "                    AND A.LSNCD = B.LSNCD "
    sStr = sStr & "                    AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                    AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                    AND A.TCRCD = '" & aTcrCD & "'"
    sStr = sStr & "                  GROUP BY A.YM, A.ACID, A.TCRCD, A.SUBJCD, A.LSNCD"
    sStr = sStr & "                 )"
    sStr = sStr & "          GROUP BY ACID, TCRCD, SUBJCD, LSNCD"
    sStr = sStr & "     )"
    sStr = sStr & "  ORDER BY SUBJCD, LSNCDNM"
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sprGwamok.MaxRows = 0
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprGwamok.MaxRows = sprGwamok.MaxRows + 1
                sprGwamok.Row = sprGwamok.MaxRows
                
                sprGwamok.Col = 1
                    sTmp = " ":     If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
    
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TSISU")) = False Then sTmp = Trim(.Fields("TSISU"))
                    If IsNumeric(sTmp) = True Then
                        If CLng(sTmp) <> 0 Then Call basFunction.Set_SprType_Numeric(sprGwamok, 0, -9999, 9999, "", CLng(sTmp))
                    End If
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SISU")) = False Then sTmp = Trim(.Fields("SISU"))
                    If IsNumeric(sTmp) = True Then
                        If CLng(sTmp) <> 0 Then Call basFunction.Set_SprType_Numeric(sprGwamok, 0, -9999, 9999, "", CLng(sTmp))
                    End If
                sprGwamok.Col = sprGwamok.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TT")) = False Then sTmp = Trim(.Fields("TT"))
                    If IsNumeric(sTmp) = True Then
                        If CLng(sTmp) <> 0 Then Call basFunction.Set_SprType_Numeric(sprGwamok, 0, -9999, 9999, "", CLng(sTmp))
                    End If
                
                .MoveNext
            Next nRec
        End If
    End With
    
    With sprGwamok
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "강의과목 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사별 과목조회"
End Sub

'## 선택 강사의 시수내역 조회
Private Sub Disp_Tcr_Sisu(ByVal aAcID As String, ByVal aTcrCD As String)
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRow        As Long
    Dim nCol        As Long
    
    On Error GoTo ErrStmt
    
'## clear
    With sprSisu
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
    End With

    sStr = ""
    sStr = sStr & "    SELECT LESSON, WEEKS,"
    sStr = sStr & "           CASE WHEN PRT_LSN = '000' THEN"
    
    Select Case Trim(basModule.SchCD)
        Case "N", "J"
            sStr = sStr & "        SUBSTR(KAEYOL,2,1)||LSNCDNM"
        Case "S"
            sStr = sStr & "        SUBSTR(KAEYOL,2,1)||LSNCDNM"
        Case "K"
            sStr = sStr & "        LSNCDNM_K"
    End Select
    
    sStr = sStr & "           ELSE"
    sStr = sStr & "               PRT_LSN"
    sStr = sStr & "           END AS LSN"
    
    sStr = sStr & "      FROM ("
    
    sStr = sStr & "            /* 정규반 내용 */ "
    sStr = sStr & "            SELECT A.ACID, A.LSNCD, A.TCRCD, B.KAEYOL, B.LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, '000' AS PRT_LSN, "
    sStr = sStr & "                   SUBSTR(GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD), 1, 1)||B.LSNCDNM AS LSNCDNM_K "
    sStr = sStr & "              From SDTRX50TB A, SDLSN01TB B"
    sStr = sStr & "             WHERE A.ACID = B.ACID  "
    sStr = sStr & "               AND A.LSNCD= B.LSNCD "
    sStr = sStr & "               AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "               AND A.ACID = '" & Trim(aAcID) & "'"
    sStr = sStr & "               AND A.TCRCD = '" & Trim(aTcrCD) & "'"
    
    sStr = sStr & "            UNION ALL"
    
    sStr = sStr & "            /* 이동반 내용 */ "
    sStr = sStr & "            SELECT A.ACID, A.LSNCD, A.TCRCD, B.KAEYOL, B.LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, '000' AS PRT_LSN, "
    sStr = sStr & "                   SUBSTR(GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD), 1, 1)||B.LSNCDNM AS LSNCDNM_K "
    sStr = sStr & "              From SDTRX50TB A, SDLSN02TB B"
    sStr = sStr & "             WHERE A.ACID = B.ACID "
    sStr = sStr & "               AND A.LSNCD= B.LSNCD "
    sStr = sStr & "               AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "               AND A.ACID = '" & Trim(aAcID) & "'"
    sStr = sStr & "               AND A.TCRCD = '" & Trim(aTcrCD) & "'"
    
    sStr = sStr & "            UNION ALL"
    
    sStr = sStr & "            /* 직접 입력내용 */ "
    sStr = sStr & "            SELECT A.ACID, A.LSNCD, A.TCRCD, PRT_KAEYOL AS KAEYOL, PRT_LSNNM AS LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, A.PRT_LSN, "
    sStr = sStr & "                   SUBSTR(GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD), 1, 1)||'00' AS LSNCDNM_K "
    sStr = sStr & "              FROM SDTRX50TB A "
    sStr = sStr & "             WHERE A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "               AND A.ACID  = '" & Trim(aAcID) & "'"
    sStr = sStr & "               AND A.TCRCD = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "               AND A.LSNCD = '00000'"
    sStr = sStr & "            )"
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRow = 1 To .RecordCount Step 1
                
                If IsNumeric(.Fields("LESSON")) = True And IsNumeric(.Fields("WEEKS")) = True Then
                    
                    Select Case CInt(Trim(.Fields("WEEKS")))        '< 요일
                        Case 2      '< 월요일
                            sprSisu.Col = 1
                        Case 3
                            sprSisu.Col = 2
                        Case 4
                            sprSisu.Col = 3
                        Case 5
                            sprSisu.Col = 4
                        Case 6
                            sprSisu.Col = 5
                        Case 7
                            sprSisu.Col = 6
                        Case 1      '< 일요일
                            sprSisu.Col = 7
                    End Select
                    
                    sprSisu.Row = CInt(.Fields("LESSON"))
                    
                    If Trim(sprSisu.Text) > " " Then        '<< 중복내용
                        sTmp = Trim(sprSisu.Text)
                        If IsNull(.Fields("LSN")) = False Then sTmp = sTmp & "/" & Trim(.Fields("LSN"))
                            Call basFunction.Set_SprType_Text(sprSisu, "TOP", "LEFT", LenB(sTmp), sTmp)
                            sprSisu.TypeEditMultiLine = True
                        
                        sprSisu.Row2 = sprSisu.Row:     sprSisu.Col2 = sprSisu.Col
                        sprSisu.BlockMode = True
                            sprSisu.BackColor = basModule.SectionColor1
                            sprSisu.BackColorStyle = BackColorStyleUnderGrid
                        sprSisu.BlockMode = False
                    Else                                    '<< 신규내용
                        sTmp = " ":     If IsNull(.Fields("LSN")) = False Then sTmp = Trim(.Fields("LSN"))
                            Call basFunction.Set_SprType_Text(sprSisu, "CENTER", "CENTER", LenB(sTmp), sTmp)
                        
                        sprSisu.Row2 = sprSisu.Row:     sprSisu.Col2 = sprSisu.Col
                        sprSisu.BlockMode = True
                            sprSisu.BackColor = basModule.SelectColor2
                            sprSisu.BackColorStyle = BackColorStyleUnderGrid
                        sprSisu.BlockMode = False
                    End If
                    
                End If
                
                .MoveNext
            Next nRow
        End If
    End With
    
    With sprSisu
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "강사시수 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사시수 조회"
    
End Sub

Private Sub sprSubj_Click(ByVal Col As Long, ByVal Row As Long)
     Dim sAcID       As String
    Dim sTcrCD      As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprSubj
        If .Tag = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):      .Row2 = .Row
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


















'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 스프레드 작업 ################################################################################

Private Sub sprTmr_Lsn_DblClick(ByVal Col As Long, ByVal Row As Long)
    With sprTmr_Lsn
        txtinSpr.Text = "SPRTMR_LSN"
        txtinRow.Text = Trim(CStr(Row))
        txtinCol.Text = Trim(CStr(Col))

        .Row = Row
        .Col = Col
            txtData.Text = Trim(.Text)
            txtData.SetFocus

    End With
End Sub


Private Sub sprTmr_Tcr_DblClick(ByVal Col As Long, ByVal Row As Long)
    With sprTmr_Tcr
        txtinSpr.Text = "SPRTMR_TCR"
        txtinRow.Text = Trim(CStr(Row))
        txtinCol.Text = Trim(CStr(Col))

        .Row = Row
        .Col = Col
            txtData.Text = Trim(.Text)
            txtData.SetFocus

    End With
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case UCase(Trim(txtinSpr.Text))
        Case "SPRTMR_LSN"
            With sprTmr_Lsn
                .Row = CLng(txtinRow.Text)
                .Col = CLng(txtinCol.Text)
                    Call basFunction.Set_SprType_Text(sprTmr_Lsn, "CENTER", "LEFT", 60, Trim(UCase(txtData.Text)))
                    
            End With
        Case "SPRTMR_TCR"
            With sprTmr_Tcr
                .Row = CLng(txtinRow.Text)
                .Col = CLng(txtinCol.Text)
                    Call basFunction.Set_SprType_Text(sprTmr_Tcr, "CENTER", "LEFT", 60, Trim(UCase(txtData.Text)))
                    
            End With
    End Select
        
    
End Sub




'## 선택
Private Sub sprTmr_Lsn_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sData       As String
    Dim sTmp        As String
    
    Dim sTcrCD      As String
    Dim sDiv()      As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    Dim sStr        As String
    
    Dim sKLs        As String
    
    Dim ni          As Long
    
    On Error Resume Next
    
    cmdSave_LSN.Enabled = True
    cmdSave_Tcr.Enabled = False
    
    With sprTmr_Lsn
        
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
'        '<<---------------------------------------------
'        txtinSpr.Text = "SPRTMR_LSN"
'        txtinRow.Text = Trim(CStr(Row))
'        txtinCol.Text = Trim(CStr(Col))
'
'        .Row = Row
'        .Col = Col
'            txtData.Text = Trim(.Text)
'            txtData.SetFocus
'        '--------------------------------------------->>
        
        
        .Col = Col
        .Row = SpreadHeader + 3:        sTmp = Right(Trim(.Text), 1)
        .Row = SpreadHeader + 2:        sTmp = sTmp & Trim(.Text)
            sKLs = sTmp
            
        sTmp = ""
        
        .Row = Row
            sData = Trim(.Text)
        
        If sData = "" Then Exit Sub
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    sTmp = Trim(.Text)
                
                '<< duplication >>
                If InStr(1, sTmp, "/", vbTextCompare) > 0 Then
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.SectionColor1
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                End If
                
                If StrComp(sData, sTmp, vbTextCompare) = 0 Then
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.SelectColor1
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                End If
            Next nCol
        Next nRow
    End With
    
    If InStr(1, sData, ",", vbTextCompare) > 0 Then
        sDiv = Split(sData, ",", -1, vbTextCompare)
        
        If UBound(sDiv) >= 1 Then
            sStr = ""
            sStr = sStr & " SELECT TCRCD "
            sStr = sStr & "   From SDTCR01TB"
            sStr = sStr & "  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND TRIM(TCRNM)  = '" & Trim(sDiv(1)) & "'"
            sStr = sStr & "    AND TRIM(SUBJNM) = '" & Trim(sDiv(0)) & "'"
            sStr = sStr & "  GROUP BY TCRCD "
            
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            


            
            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
        
            With DBRec
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    sTcrCD = "":    If IsNull(.Fields("TCRCD")) = False Then sTcrCD = Trim(.Fields("TCRCD"))
                    
                    Call Sel_TmrTcr_Data(sKLs)
                    Call Sel_SprTCR(sTcrCD)
                    
                    'Call Disp_Tcr_Sisu(Trim(basModule.SchCD), sTcrCD)       '< 선택 강사의 시수내역 조회
                End If
            End With
        End If
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    
End Sub


'## 선택
Private Sub Sel_TmrTcr_Data(ByVal aKLs As String)
    Dim nRow        As Long
    Dim nCol        As Long
    
    With sprTmr_Tcr
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                If InStr(1, Trim(.Text), aKLs, vbTextCompare) > 0 Then
                    .Row2 = .Row:   .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.SelectColor1
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .SetActiveCell .Col, .Row
                Else
                    '<< duplication >>
                    If InStr(1, Trim(.Text), "/", vbTextCompare) > 0 Then
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.SectionColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
                    If .BackColor = basModule.SectionColor1 Or _
                       .BackColor = lblNotTeaching.BackColor Then
                        ' no action
                    Else
                        .Row2 = .Row:   .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                
                End If
                    
            Next nCol
        Next nRow
    End With
End Sub


Private Sub sprTmr_Tcr_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sData       As String
    Dim sTmp        As String
    
    Dim sTcrCD      As String
    
    Dim sTcrNM      As String
    
    cmdSave_LSN.Enabled = False
    cmdSave_Tcr.Enabled = True
    
    With sprTmr_Tcr
        
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
'        '<<----------------------------------------
'        txtinSpr.Text = "SPRTMR_TCR"
'        txtinRow.Text = Trim(CStr(Row))
'        txtinCol.Text = Trim(CStr(Col))
'
'        .Row = Row
'        .Col = Col
'            txtData.Text = Trim(.Text)
'            txtData.SetFocus
'        '---------------------------------------->>
        
        .Row = Row
        .Col = Col
            sData = Trim(.Text)
        .Col = SpreadHeader
            sTcrCD = Trim(.Text)
        .Col = SpreadHeader + 2
            sTcrNM = Trim(.Text)
        
        If sData = "" Then
            '< 초기화
            For nRow = 1 To sprSisu.MaxRows Step 1
                For nCol = 1 To sprSisu.MaxCols Step 1
                    sprSisu.Row = nRow
                    sprSisu.Col = nCol
                        sprSisu.Text = ""
                Next nCol
            Next nRow
            
            sprSisu.Row = 1:   sprSisu.Row2 = sprSisu.MaxRows
            sprSisu.Col = 1:   sprSisu.Col2 = sprSisu.MaxCols
            sprSisu.BlockMode = True
                sprSisu.BackColor = basModule.WhiteColor
                sprSisu.BackColorStyle = BackColorStyleUnderGrid
            sprSisu.BlockMode = False
        
            Exit Sub
        End If
        
'        .Row = 1:   .Row2 = .MaxRows
'        .Col = 1:   .Col2 = .MaxCols
'        .BlockMode = True
'            .BackColor = basModule.WhiteColor
'            .BackColorStyle = BackColorStyleUnderGrid
'        .BlockMode = False
        
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    sTmp = Trim(.Text)
                
                Select Case .BackColor
                
                    Case basModule.SectionColor1
                        ' no action
                    Case lblNotTeaching.BackColor
                        ' no action
                    Case Else
                        
                        If Trim(.Text) > " " Then
                            .Row2 = .Row
                            .Col2 = .Col
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                        End If
                        
                        '<< duplication >>
                        If InStr(1, Trim(.Text), "/", vbTextCompare) > 0 Then
                            .Row2 = .Row
                            .Col2 = .Col
                            .BlockMode = True
                                .BackColor = basModule.SectionColor1
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                        End If
                        
                        If StrComp(sData, sTmp, vbTextCompare) = 0 Then
                            .Row2 = .Row
                            .Col2 = .Col
                            .BlockMode = True
                                .BackColor = basModule.SelectColor1
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                        End If
                        
                End Select
            Next nCol
        Next nRow
        
        '<< 강사 선택 >>
        'sTcrCD
        'sTcrNM
        Call Sel_LsnTcr_Data(sTcrNM)
        Call Sel_SprTCR(sTcrCD)
        
        'Call Disp_Tcr_Sisu(Trim(basModule.SchCD), sTcrCD)       '< 선택 강사의 시수내역 조회
        
    End With
End Sub

'## 강사 선택
Private Sub Sel_SprTCR(ByVal aTcrCD As String)
    Dim nRow        As Long
    Dim nCol        As Long
    
    With sprTcr
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 2
            If StrComp(aTcrCD, Trim(.Text), vbTextCompare) = 0 Then
                .SetActiveCell 3, .Row
                Call sprTcr_Click(3, .Row)
                
                Exit Sub
            End If
        Next nRow
    End With
End Sub

Private Sub Sel_LsnTcr_Data(ByVal aTcrNM As String)
    Dim nRow        As Long
    Dim nCol        As Long
    
    With sprTmr_Lsn
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                If InStr(1, Trim(.Text), aTcrNM, vbTextCompare) > 0 Then
                    .Row2 = .Row:   .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.SelectColor1
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .SetActiveCell .Col, .Row
                    
                Else
                    '<< duplication >>
                    If InStr(1, Trim(.Text), "/", vbTextCompare) > 0 Then
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.SectionColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
                    If .BackColor = basModule.SectionColor1 Or _
                       .BackColor = lblNotTeaching.BackColor Then
                        ' no action
                    Else
                        .Row2 = .Row:   .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                
                End If
                    
            Next nCol
        Next nRow
    End With
    
End Sub

'==============================================================================================
'## 시간표 등록내역 삭제 - << 반별 내역에서 >>
'==============================================================================================
Private Sub sprTmr_Lsn_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sLsnNM      As String
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim sTmp        As String
    
    With sprTmr_Lsn
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = .ActiveCol
                    If StrComp(Trim(.Text), "X", vbTextCompare) = 0 Then Exit Sub
            
                .Row = .ActiveRow
                    .Col = SpreadHeader:        sWeek = Trim(.Text)
                    .Col = SpreadHeader + 2:    sLesson = Trim(.Text)
                .Col = .ActiveCol
                    .Row = SpreadHeader:        sLsnNM = Trim(.Text)
                
                sTmp = ""
                sTmp = sTmp & "반    【 " & sLsnNM & " 】" & vbCrLf
                sTmp = sTmp & "요일 【 " & sWeek & " 】" & vbCrLf
                sTmp = sTmp & "교시 【 " & sLesson & " 】" & vbCrLf
                sTmp = sTmp & "내용을 삭제하시겠습니까?"
                
                If MsgBox(sTmp, vbQuestion + vbYesNo, "시간삭제") = vbNo Then
                    Exit Sub
                End If
                
                .Row = .ActiveRow
                .Col = .ActiveCol
                    Call LSN_Time_Delete(.Row, .Col)        '< 삭제할 cell
                
            Case vbKeyBack
'                .Row = .ActiveRow
'                .Col = .ActiveCol
'                    .Text = ""
                
        End Select
        
        .SetFocus
        Call .SetActiveCell(.ActiveCol, .ActiveRow)
        
        
        
    End With
End Sub

Private Sub LSN_Time_Delete(ByVal aRow As Long, ByVal aCol As Long)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    Dim nExe        As Integer
    
    Dim sTmp        As String
    Dim sDiv()      As String
    
    Dim bChk        As Boolean
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sWeek       As String
    Dim sLesson     As String
    Dim sLsnCD      As String
    
    Dim sStr        As String
    
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    On Error GoTo ErrStmt
    
    '<< 삭제대상 강사와 과목내역 찾기 >>
    bChk = False
    
    sprTmr_Lsn.Row = aRow
    sprTmr_Lsn.Col = aCol
        sTmp = Trim(sprTmr_Lsn.Text)
    
    If InStr(1, sTmp, ",", vbTextCompare) > 0 Then
        sDiv = Split(sTmp, ",", -1, vbTextCompare)
        
        sStr = ""
        sStr = sStr & " SELECT TCRCD, SUBJCD"
        sStr = sStr & "   From SDTCR01TB"
        sStr = sStr & "  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "    AND TRIM(TCRNM)  = '" & Trim(sDiv(1)) & "'"
        sStr = sStr & "    AND TRIM(SUBJNM) = '" & Trim(sDiv(0)) & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
     
     

        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
    
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                sTcrCD = "":    If IsNull(.Fields("TCRCD")) = False Then sTcrCD = Trim(.Fields("TCRCD"))
                sSubjCD = "":   If IsNull(.Fields("SUBJCD")) = False Then sSubjCD = Trim(.Fields("SUBJCD"))
                
                If sTcrCD > "" And sSubjCD > "" Then
                    
                    sprTmr_Lsn.Row = aRow
                        sprTmr_Lsn.Col = SpreadHeader + 1:      sWeek = Trim(sprTmr_Lsn.Text)
                        sprTmr_Lsn.Col = SpreadHeader + 2:      sLesson = Trim(sprTmr_Lsn.Text)
                        
                    sprTmr_Lsn.Col = aCol
                        sprTmr_Lsn.Row = SpreadHeader + 1:      sLsnCD = Trim(sprTmr_Lsn.Text)
                    
                    bChk = True
                    
                End If
            End If
        End With
    End If
    
    If bChk = False Then
        Set DBCmd = Nothing
        Set DBRec = Nothing
        Set DBParam = Nothing
        
        MsgBox "삭제할 수 없습니다." & vbCrLf & "관리자에게 문의하십시요.", vbCritical + vbOKOnly, "시간삭제"
        Exit Sub
    End If
    


    
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDTRX50TB "
    sStr = sStr & "  WHERE YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND TCRCD  = '" & sTcrCD & "'"
    sStr = sStr & "    AND SUBJCD = '" & sSubjCD & "'"
    sStr = sStr & "    AND LSNCD  = '" & sLsnCD & "'"
    sStr = sStr & "    AND WEEKS  = " & sWeek
    sStr = sStr & "    AND LESSON = " & sLesson
    
    basDataBase.DBConn.BeginTrans
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    If nExe >= 1 Then
        basDataBase.DBConn.CommitTrans
    Else
        basDataBase.DBConn.RollbackTrans
    End If
         
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    With sprTmr_Lsn
        .Row = aRow
        .Col = aCol
            .Text = ""                      '< 삭제처리 함.
        
        '3. 등록된 내용 저장
        '변경내용 display
        
        'sWeek
        'sLesson
        'sLsnCD
        '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
        If sTcrCD <> "" And sSubjCD <> "" Then
            Call Show_TMR_Tcr_Week_Lesson(sWeek, sLesson)
            Call sprTmr_Lsn_Click(aCol, aRow)
            
        End If
         
        ' 강사 시수내역 조정하기
        sprGwamok.MaxRows = 0                           '< 강사, 과목내역 초기화
        For nRow = 1 To sprSisu.MaxRows Step 1          '< 요일별 내역 초기화
            For nCol = 1 To sprSisu.MaxCols Step 1
                sprSisu.Row = nRow
                sprSisu.Col = nCol
                    sprSisu.Text = ""
            Next nCol
        Next nRow

        sprSisu.Row = 1:        sprSisu.Row2 = sprSisu.MaxRows
        sprSisu.Col = 1:        sprSisu.Col2 = sprSisu.MaxCols
        sprSisu.BlockMode = True
            sprSisu.BackColor = basModule.WhiteColor
            sprSisu.BackColorStyle = BackColorStyleUnderGrid
        sprSisu.BlockMode = False

        For nRow = 1 To sprTcr.MaxRows Step 1           '< 강사시수처리
            sprTcr.Row = nRow
            sprTcr.Col = 2
            If StrComp(sTcrCD, Trim(sprTcr.Text), vbTextCompare) = 0 Then
                Select Case sWeek
                    Case "2"
                        sprTcr.Col = 6
                        fpT(2).Value = fpT(2).Value - 1
                    Case "3"
                        sprTcr.Col = 7
                        fpT(3).Value = fpT(3).Value - 1
                    Case "4"
                        sprTcr.Col = 8
                        fpT(4).Value = fpT(4).Value - 1
                    Case "5"
                        sprTcr.Col = 9
                        fpT(5).Value = fpT(5).Value - 1
                    Case "6"
                        sprTcr.Col = 10
                        fpT(6).Value = fpT(6).Value - 1
                    Case "7"
                        sprTcr.Col = 11
                        fpT(7).Value = fpT(7).Value - 1
                    Case "1"
                        sprTcr.Col = 12
                        
                End Select

                If Trim(sprTcr.Text) = "" Then
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", 0)
                Else
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) - 1)
                End If

                sprTcr.Col = 5
                    fpT(1).Value = fpT(1).Value + 1
                If (sprTcr.Text) = "" Then
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", 1)
                Else
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) + 1)
                End If

            End If
        Next nRow

        If Trim(sTcrCD) <> "" Then
            Call Sel_SprTCR(sTcrCD)
        End If

        .Row2 = .Row
        .Col2 = .Col
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
    End With
    
    
    MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간삭제"
    
    Exit Sub
    
ErrStmt:
    On Error GoTo 0
    On Error Resume Next
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    MsgBox "시간삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간삭제"
    On Error GoTo 0
    
End Sub


'## 전체 시간표 내역에서 보여주기
Public Sub Show_TMR_Tcr_Week_Lesson(ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRec        As Long
    Dim ni          As Long
    Dim sData       As String

    Dim nRow        As Long
    Dim nCol        As Long

    
    Dim sLesson     As String
    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sTmpTcrCD   As String
    Dim sTmpSubjCD  As String
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long

    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD, GET_KEAYOL_N_LSN_TCR01(A.ACID, A.LSNCD) AS DS"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                             '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    
    With sprTmr_Tcr
        For nCol = 1 To .MaxCols Step 1
            
            .Col = nCol:        nChkCol = .Col
                .Row = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
                .Row = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
                
            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = nChkCol
                        .Text = ""
                        
                    If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                       sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                        ' no action
                    Else
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
                Next nRow
            End If
        Next nCol
    End With


    DBRec.MoveFirst
    For nRec = 1 To DBRec.RecordCount Step 1
        
        If IsNull(DBRec.Fields("TCRCD")) = False And _
           IsNull(DBRec.Fields("SUBJCD")) = False And _
           IsNull(DBRec.Fields("DS")) = False Then
            
            sTcrCD = Trim(DBRec.Fields("TCRCD"))
            sSubjCD = Trim(DBRec.Fields("SUBJCD"))
            sData = Trim(DBRec.Fields("DS"))
            
            With sprTmr_Tcr
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow:        nChkRow = .Row
                        .Col = SpreadHeader:            sTmpTcrCD = Trim(.Text)
                        .Col = SpreadHeader + 1:        sTmpSubjCD = Trim(.Text)
                    
                    If StrComp(sTcrCD, sTmpTcrCD, vbTextCompare) = 0 And _
                       StrComp(sSubjCD, sTmpSubjCD, vbTextCompare) = 0 Then
                    
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol:    nChkCol = .Col
                                .Row = SpreadHeader + 1:    sTmpWeek = Trim(.Text)
                                .Row = SpreadHeader + 2:    sTmpLesson = Trim(.Text)
                                
                            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                            
                                .Row = nChkRow
                                .Col = nChkCol
                                
                                If Trim(.Text) = "" Then
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                        
                                        If InStr(1, sData, "/", vbTextCompare) > 0 Then
                                            .Row2 = .Row
                                            .Col2 = .Col
                                            .BlockMode = True
                                                .BackColor = basModule.SectionColor1
                                                .BackColorStyle = BackColorStyleUnderGrid
                                            .BlockMode = False
                                        End If
                                    End If
                                End If
                            End If
                        Next nCol
                    End If
                Next nRow
                
            End With
        End If
        
        DBRec.MoveNext
    Next nRec
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
        
End Sub


















'## 시간표 등록내역 삭제 - << 강사별 내역에서 >>
Private Sub sprTmr_Tcr_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTcrNM      As String
    Dim sSubjNM     As String
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim sTmp        As String
    
    Dim ni          As Long
    Dim nRow        As Long
    Dim nCol        As Long
    
    ReDim uTcr_Dup_Row_and_Col(0) As tTcr_Dup_Row_and_Col
    
    With sprTmr_Tcr
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub
        
        ni = 0
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                
                If .BackColor = basModule.SectionColor1 Then
                    ni = ni + 1
                    
                    ReDim Preserve uTcr_Dup_Row_and_Col(ni) As tTcr_Dup_Row_and_Col
                    uTcr_Dup_Row_and_Col(ni).Row = .Row
                    uTcr_Dup_Row_and_Col(ni).Col = .Col
                    
                End If
            Next nCol
        Next nRow
        
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = .ActiveCol
                    If StrComp(Trim(.Text), "X", vbTextCompare) = 0 Then Exit Sub
            
                .Row = .ActiveRow
                    .Col = SpreadHeader + 2:    sTcrNM = Trim(.Text)
                    .Col = SpreadHeader + 3:    sSubjNM = Trim(.Text)
                    
                .Col = .ActiveCol
                    .Row = SpreadHeader:        sWeek = Trim(.Text)
                    .Row = SpreadHeader + 2:    sLesson = Trim(.Text)
                    
                sTmp = ""
                sTmp = sTmp & "강사 【 " & sTcrNM & " 】" & vbCrLf
                sTmp = sTmp & "과목 【 " & sSubjNM & " 】" & vbCrLf
                sTmp = sTmp & "요일 【 " & sWeek & " 】" & vbCrLf
                sTmp = sTmp & "교시 【 " & sLesson & " 】" & vbCrLf
                sTmp = sTmp & "의 내용을 삭제하시겠습니까?"
                
                If MsgBox(sTmp, vbQuestion + vbYesNo, "시간삭제") = vbNo Then
                    Exit Sub
                End If
                
                .Row = .ActiveRow
                .Col = .ActiveCol
                    Call TCR_Time_Delete(.Row, .Col)        '< 삭제대상
                
        End Select
        
        .SetFocus
        .SetActiveCell .ActiveCol, .ActiveRow
        
    End With
End Sub

Private Sub TCR_Time_Delete(ByVal aRow As Long, ByVal aCol As Long)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    Dim nExe        As Integer
    Dim sStr        As String
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sWeek       As String
    Dim sLesson     As String
    Dim sLsnCD      As String
    
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    On Error GoTo ErrStmt
    
    '<< 삭제대상 강사와 과목내역 찾기 >>
    
    With sprTmr_Tcr
        .Row = .ActiveRow
            .Col = SpreadHeader:            sTcrCD = Trim(.Text)
            .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
            
        .Col = .ActiveCol
            .Row = SpreadHeader + 1:        sWeek = Trim(.Text)
            .Row = SpreadHeader + 2:        sLesson = Trim(.Text)
    End With
    
    
        
    sStr = ""
    sStr = sStr & "        SELECT ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS"
    sStr = sStr & "          FROM (SELECT A.ACID, B.LSNCD, A.TCRCD, B.KAEYOL, B.LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, A.PRT_LSN, A.SUBJCD"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, A.TCRCD, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, '000' AS PRT_LSN, A.SUBJCD"
    sStr = sStr & "                          From SDTRX50TB A, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) B"

    sStr = sStr & "                         WHERE A.ACID = B.ACID "
    sStr = sStr & "                           AND A.LSNCD= B.LSNCD"
    sStr = sStr & "                           AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                          From SDLSN01TB"
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        UNION"                                                   '2009.01.12 추가
    sStr = sStr & "                        SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                          FROM SDLSN02TB "
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) B"
    sStr = sStr & "                 Where A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                   AND A.TCRCD  = '" & sTcrCD & "'"
    sStr = sStr & "                   AND A.SUBJCD = '" & sSubjCD & "'"
    sStr = sStr & "                   AND A.WEEKS  = " & sWeek
    sStr = sStr & "                   AND A.LESSON = " & sLesson
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.ACID, B.LSNCD, A.TCRCD, B.KAEYOL, B.LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, A.PRT_LSN, A.SUBJCD"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, A.TCRCD, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, '000' AS PRT_LSN, A.SUBJCD"
    sStr = sStr & "                          From SDTRX50TB A, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) B"
    
    sStr = sStr & "                         WHERE A.ACID = B.ACID  "
    sStr = sStr & "                           AND A.LSNCD= B.LSNCD "
    sStr = sStr & "                           AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                          From SDLSN02TB"
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        ) B"
    sStr = sStr & "                 Where A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                   AND A.TCRCD  = '" & sTcrCD & "'"
    sStr = sStr & "                   AND A.SUBJCD = '" & sSubjCD & "'"
    sStr = sStr & "                   AND A.WEEKS  = " & sWeek
    sStr = sStr & "                   AND A.LESSON = " & sLesson
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.ACID, A.LSNCD, A.TCRCD, '' AS KAEYOL, '00' AS LSNCDNM, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM, A.LESSON, A.WEEKS, A.PRT_LSN, A.SUBJCD"
    sStr = sStr & "                  FROM SDTRX50TB A, "
    
    sStr = sStr & "                       (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                              '2009.01.12 추가
    sStr = sStr & "                          FROM SDLSN01TB "
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        UNION"
    sStr = sStr & "                        SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                          FROM SDLSN02TB "
    sStr = sStr & "                         WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       ) B"

    sStr = sStr & "                 WHERE A.ACID   = B.ACID "
    sStr = sStr & "                   AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                   AND A.TCRCD  = '" & sTcrCD & "'"
    sStr = sStr & "                   AND A.SUBJCD = '" & sSubjCD & "'"
    sStr = sStr & "                   AND A.WEEKS  = " & sWeek
    sStr = sStr & "                   AND A.LESSON = " & sLesson
    sStr = sStr & "                )"
    sStr = sStr & "          GROUP BY ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS"
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
        


    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
        
    With DBRec
        If .RecordCount = 0 Then
            With sprTmr_Tcr
                .Row = aRow
                .Col = aCol
                    .Text = ""
                
                If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                   sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                    ' no action
                Else
                    .Row = 1:       .Row2 = .MaxRows
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                End If
            End With
            
        ElseIf .RecordCount >= 1 Then
            .MoveFirst
            
            
            sLsnCD = "":    If IsNull(.Fields("LSNCD")) = False Then sLsnCD = Trim(.Fields("LSNCD"))
            If sLsnCD > "" Then
                
                '<< 삭제 데이터 모두 조합 >>
                
                'TCRCD
                'SUBJCD
                'WEEK
                'LESSON
                'LSNCD
                


                
                sStr = ""
                sStr = sStr & " DELETE "
                sStr = sStr & "   FROM SDTRX50TB "
                sStr = sStr & "  WHERE YM     = '" & Trim(fpYM.UnFmtText) & "'"
                sStr = sStr & "    AND ACID   = '" & Trim(basModule.SchCD) & "'"
                sStr = sStr & "    AND TCRCD  = '" & sTcrCD & "'"
                sStr = sStr & "    AND SUBJCD = '" & sSubjCD & "'"
                sStr = sStr & "    AND LSNCD  = '" & sLsnCD & "'"
                sStr = sStr & "    AND WEEKS  = " & sWeek
                sStr = sStr & "    AND LESSON = " & sLesson
                
                basDataBase.DBConn.BeginTrans
    
                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
                
                DBCmd.Execute nExe, , -1
                If nExe >= 1 Then
                    basDataBase.DBConn.CommitTrans
                Else
                    basDataBase.DBConn.RollbackTrans
                End If
                
            End If
        Else
            Set DBCmd = Nothing
            Set DBRec = Nothing
            Set DBParam = Nothing
        
            MsgBox "삭제할 수 없습니다." & vbCrLf & "관리자에게 문의하십시요.", vbCritical + vbOKOnly, "시간삭제"
            Exit Sub
        End If
    End With
     
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    
    With sprTmr_Tcr
        .Row = aRow
        .Col = aCol
            .Text = ""
        
        '<< 강사시수 조정
        .Row = aRow
        .Col = SpreadHeader + 6
        If Trim(.Text) = "" Then
            .Text = "1"
        Else
            .Text = Trim(CStr(CLng(.Text) + 1))
        End If
        
        '< 내역 보여주기
        Call Show_TMR_Lsn_Week_Lesson(sWeek, sLesson)
        
        ' 강사 시수내역 조정하기
        sprGwamok.MaxRows = 0                           '< 강사, 과목내역 초기화
        For nRow = 1 To sprSisu.MaxRows Step 1          '< 요일별 내역 초기화
            For nCol = 1 To sprSisu.MaxCols Step 1
                sprSisu.Row = nRow
                sprSisu.Col = nCol
                    sprSisu.Text = ""
            Next nCol
        Next nRow
        
        sprSisu.Row = 1:        sprSisu.Row2 = sprSisu.MaxRows
        sprSisu.Col = 1:        sprSisu.Col2 = sprSisu.MaxCols
        sprSisu.BlockMode = True
            sprSisu.BackColor = basModule.WhiteColor
            sprSisu.BackColorStyle = BackColorStyleUnderGrid
        sprSisu.BlockMode = False
        
        For nRow = 1 To sprTcr.MaxRows Step 1           '< 강사시수처리
            sprTcr.Row = nRow
            sprTcr.Col = 2
            If StrComp(sTcrCD, Trim(sprTcr.Text), vbTextCompare) = 0 Then
                Select Case sWeek
                    Case "2"
                        sprTcr.Col = 6
                        fpT(2).Value = fpT(2).Value - 1
                    Case "3"
                        sprTcr.Col = 7
                        fpT(3).Value = fpT(3).Value - 1
                    Case "4"
                        sprTcr.Col = 8
                        fpT(4).Value = fpT(4).Value - 1
                    Case "5"
                        sprTcr.Col = 9
                        fpT(5).Value = fpT(5).Value - 1
                    Case "6"
                        sprTcr.Col = 10
                        fpT(6).Value = fpT(6).Value - 1
                    Case "7"
                        sprTcr.Col = 11
                        fpT(7).Value = fpT(7).Value - 1
                    Case "1"
                        sprTcr.Col = 12
                End Select
                
                If Trim(sprTcr.Text) = "" Then
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", 0)
                Else
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) - 1)
                End If
                
                sprTcr.Col = 5
                    fpT(1).Value = fpT(1).Value + 1
                If Trim(sprTcr.Text) = "" Then
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", 1)
                Else
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) + 1)
                End If
                
            End If
        Next nRow
        
        
        .Row = aRow:       .Row2 = .Row
        .Col = aCol:       .Col2 = .Col
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    
    '<< duplicate 되어진 항목 표시 >>
        For ni = 1 To UBound(uTcr_Dup_Row_and_Col) Step 1
            .Row = uTcr_Dup_Row_and_Col(ni).Row
            .Col = uTcr_Dup_Row_and_Col(ni).Col
            
            If Trim(.Text) <> "" Then
                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = basModule.SectionColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End If
        Next ni
        
        If Trim(sTcrCD) <> "" Then
            Call Sel_SprTCR(sTcrCD)
        End If
    End With
    
    MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간삭제"
    
    Exit Sub
    
ErrStmt:
    On Error GoTo 0
    On Error Resume Next
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    MsgBox "시간삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간삭제"
    On Error GoTo 0
End Sub


'## 전체 시간표 내역에서 보여주기
Public Sub Show_TMR_Lsn_Week_Lesson(ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRec        As Long
    Dim ni          As Long
    Dim sData       As String

    Dim nRow        As Long
    Dim nCol        As Long

    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim sLsnCD      As String
    Dim sTmpLsnCD   As String
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long

    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD)||','||GET_TCRNM(A.ACID, A.TCRCD) AS DS, A.LSNCD, A.WEEKS, A.LESSON"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    
    With sprTmr_Lsn
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow:    nChkRow = .Row
            
            .Col = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
            .Col = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
            
            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                .Row = nChkRow
                
                For nCol = 1 To .MaxCols Step 1
                    .Col = nCol
                    .Text = ""
                Next nCol
            End If
        Next nRow
    End With
    
    
    DBRec.MoveFirst
    For nRec = 1 To DBRec.RecordCount Step 1
        
        If IsNull(DBRec.Fields("LSNCD")) = False And _
           IsNull(DBRec.Fields("DS")) = False Then
            
            sLsnCD = Trim(DBRec.Fields("LSNCD"))
            sData = Trim(DBRec.Fields("DS"))
            
            With sprTmr_Lsn
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow:        nChkRow = .Row
                        .Col = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
                        .Col = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
                    
                    If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                       StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                    
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol:    nChkCol = .Col
                                .Row = SpreadHeader + 1:    sTmpLsnCD = Trim(.Text)
                                
                            If StrComp(sLsnCD, sTmpLsnCD, vbTextCompare) = 0 Then
                                .Row = nChkRow
                                .Col = nChkCol
                                
                                If Trim(.Text) = "" Then
                                
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        Call basFunction.Set_SprType_Text(sprTmr_Lsn, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(sprTmr_Lsn, "center", "left", 60, sData)
                                        
                                        If InStr(1, sData, "/", vbTextCompare) > 0 Then
                                            .Row2 = .Row
                                            .Col2 = .Col
                                            .BlockMode = True
                                                .BackColor = basModule.SectionColor1
                                                .BackColorStyle = BackColorStyleUnderGrid
                                            .BlockMode = False
                                        End If
                                    End If
                                End If
                            End If
                        Next nCol
                    End If
                Next nRow
                
            End With
        End If
        
        DBRec.MoveNext
    Next nRec
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
        
End Sub





























'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 등록 작업 : 반 내역으로 저장 ################################################################################

'## 각 데이터 저장내역을 살펴보고, 바뀐내역 또는 없는 내역은 신규 또는 갱신함.
'## 시간은 걸리겠지만, 등록하는 시점에서 저장하는 것이 해당 프로그램에선 불가능


Private Sub cmdSave_LSN_Click()
    
    Dim nRow            As Long
    Dim nCol            As Long
    
    Dim sTmp            As String
    
    Dim sDivSlash()     As String
    Dim nRecSlash       As Long
    
    Dim sDivComma()     As String
    Dim nRecComma       As Long
    
    Dim nChkRow         As Long
    Dim nChkCol         As Long
    
    Dim sSubjNM         As String
    Dim sTcrNM          As String
    
    Dim sLsnCD          As String
    Dim sWeek           As String
    Dim sLesson         As String
    
    Dim sTcrCD          As String
    Dim sSubjCD         As String
    
    Dim nLastSaveChk    As Long
    
    cmdSave_LSN.Enabled = False
    
    With sprTmr_Lsn
        
        ProgressBar1.Min = 0
        ProgressBar1.Max = 100
        ProgressBar1.Value = 0
        
        For nRow = 1 To .MaxRows Step 1
        
            ProgressBar1.Value = Fix(nRow / .MaxRows * 100)
            
            For nCol = 1 To .MaxCols Step 1
            
                .Row = nRow:        nChkRow = .Row
                .Col = nCol:        nChkCol = .Col
                
                If StrComp(Trim(.Text), "X", vbTextCompare) <> 0 Then       ' X값으로 들어오면 반려
                
                '===============================================================
                '## 일반적 데이터
                '===============================================================
                    If InStr(1, Trim(.Text), "/", vbTextCompare) = 0 Then
                        
                        If InStr(1, Trim(.Text), ",", vbTextCompare) > 0 Then       ' 반드시 【 , 】 로 과목 및 강사로 구분
                            sDivComma() = Split(UCase(Trim(.Text)), ",", -1, vbTextCompare)
                            
                            sSubjNM = Trim(sDivComma(0))        '< 과목명
                            sTcrNM = Trim(sDivComma(1))         '< 강사명
                            
                            .Row = nChkRow
                                .Col = SpreadHeader + 1:        sWeek = Trim(.Text)
                                .Col = SpreadHeader + 2:        sLesson = Trim(.Text)
                            .Col = nChkCol
                                .Row = SpreadHeader + 1:        sLsnCD = Trim(.Text)
                                
                            '<< 과목 및 강사 조회 >>
                            If sLsnCD = "00000" Then
                                MsgBox "현재 등록한 반의 내용은 입력할 수 없는 작업입니다." & vbCrLf & _
                                       "작업을 위해선 아래의 시간표에서" & vbCrLf & _
                                       "등록방법 -> 1,2 계열제외는 X01(3), 계열(1), 표시반명(10) 넣으세요. 입력후 반드시 엔터키를 치십시요.", vbExclamation + vbOKOnly, "시간표 등록"
                                       
                                cmdSave_LSN.Enabled = True
                                GoTo GONEXT     '< 다음 진행
                            End If
                            
                            sTcrCD = "":        sSubjCD = ""
                            Call Find_Tcr_and_Subj_Code(sTcrCD, sSubjCD, sTcrNM, sSubjNM)
                            
                        '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
                            If sTcrCD <> "" And sSubjCD <> "" Then
                            
                            '1. 기존 등록된 내용을 살펴본다.
                            '   단, 현재 자기의 필드에 있는 내용은 무시
                                nLastSaveChk = 0
                                nLastSaveChk = Find_Already_Save_TCR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                If nLastSaveChk > 0 Then
                                    sTmp = ""
                                    Select Case sWeek
                                        Case "2"
                                            sTmp = sTmp & "월"
                                        Case "3"
                                            sTmp = sTmp & "화"
                                        Case "4"
                                            sTmp = sTmp & "수"
                                        Case "5"
                                            sTmp = sTmp & "목"
                                        Case "6"
                                            sTmp = sTmp & "금"
                                        Case "7"
                                            sTmp = sTmp & "토"
                                        Case "1"
                                            sTmp = sTmp & "일"
                                    End Select
                                    sTmp = sTmp & "요일 " & sLesson & "교시에서 【다른반에 강의】를 합니다." & vbCrLf & "등록하시겠습니까?"
                                    
'                                    If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                        cmdSave_LSN.Enabled = True
'                                        Exit Sub
'                                    End If

                                    GoTo GONEXT     '< 다음 진행
                                    
                                End If
                                
                            '2. 기존 등록된 내용을 살펴본다.
                            '   단, 현재 자기의 필드에 있는 내용은 무시
                                nLastSaveChk = 0
                                nLastSaveChk = Find_Already_Save_LSN_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                If nLastSaveChk > 0 Then
                                    sTmp = ""
                                    Select Case sWeek
                                        Case "2"
                                            sTmp = sTmp & "월"
                                        Case "3"
                                            sTmp = sTmp & "화"
                                        Case "4"
                                            sTmp = sTmp & "수"
                                        Case "5"
                                            sTmp = sTmp & "목"
                                        Case "6"
                                            sTmp = sTmp & "금"
                                        Case "7"
                                            sTmp = sTmp & "토"
                                        Case "1"
                                            sTmp = sTmp & "일"
                                    End Select
                                    sTmp = sTmp & "요일 " & sLesson & "교시에서 【같은 강의실에서 강의하는 강사】가 있습니다." & vbCrLf & "등록하시겠습니까?"
                                    
'                                    If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                        cmdSave_LSN.Enabled = True
'                                        Exit Sub
'                                    End If

                                    GoTo GONEXT     '< 다음 진행
                                    
                                End If
                                
                                
                            '** 시간표 내역 등록하기 **
                                Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                Call Show_TMR_Tcr(sLsnCD, sWeek, sLesson)
                            
                            End If
                        End If
                        
                '===============================================================
                '## 일반적 데이터
                '===============================================================
                    Else
                        
                        sDivSlash() = Split(Trim(.Text), "/", -1, vbTextCompare)
                        For nRecSlash = 0 To UBound(sDivSlash) - 1 Step 1
                        
                            If InStr(1, Trim(sDivSlash(nRecSlash)), ",", vbTextCompare) > 0 Then       ' 반드시 【 , 】 로 과목 및 강사로 구분
                                sDivComma() = Split(Trim(sDivSlash(nRecSlash)), ",", -1, vbTextCompare)
                                
                                sSubjNM = Trim(sDivComma(0))        '< 과목명
                                sTcrNM = Trim(sDivComma(1))         '< 강사명
                                
                                .Row = nChkRow
                                    .Col = SpreadHeader + 1:        sWeek = Trim(.Text)
                                    .Col = SpreadHeader + 2:        sLesson = Trim(.Text)
                                .Col = nChkCol
                                    .Row = SpreadHeader + 1:        sLsnCD = Trim(.Text)
                                
                                '<< 과목 및 강사 조회 >>
                                If sLsnCD = "00000" Then
                                    MsgBox "현재 등록한 반의 내용은 입력할 수 없는 작업입니다." & vbCrLf & _
                                           "작업을 위해선 아래의 시간표에서" & vbCrLf & _
                                           "등록방법 -> 1,2 계열제외는 X01(3), 계열(1), 표시반명(10) 넣으세요. 입력후 반드시 엔터키를 치십시요.", vbExclamation + vbOKOnly, "시간표 등록"
                                           
                                    cmdSave_LSN.Enabled = True
                                    GoTo GONEXT     '< 다음진행
                                End If
                                
                                sTcrCD = "":        sSubjCD = ""
                                Call Find_Tcr_and_Subj_Code(sTcrCD, sSubjCD, sTcrNM, sSubjNM)
                                
                            '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
                                If sTcrCD <> "" And sSubjCD <> "" Then
                                
                                '1. 기존 등록된 내용을 살펴본다.
                                '   단, 현재 자기의 필드에 있는 내용은 무시
                                    nLastSaveChk = 0
                                    nLastSaveChk = Find_Already_Save_TCR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                    If nLastSaveChk > 0 Then
                                        sTmp = ""
                                        Select Case sWeek
                                            Case "2"
                                                sTmp = sTmp & "월"
                                            Case "3"
                                                sTmp = sTmp & "화"
                                            Case "4"
                                                sTmp = sTmp & "수"
                                            Case "5"
                                                sTmp = sTmp & "목"
                                            Case "6"
                                                sTmp = sTmp & "금"
                                            Case "7"
                                                sTmp = sTmp & "토"
                                            Case "1"
                                                sTmp = sTmp & "일"
                                        End Select
                                        sTmp = sTmp & "요일 " & sLesson & "교시에서 【다른반에 강의】를 합니다." & vbCrLf & "등록하시겠습니까?"
                                        
'                                        If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                            cmdSave_LSN.Enabled = True
'                                            Exit Sub
'                                        End If

                                        GoTo GONEXT     '< 다음 진행
                                        
                                    End If
                                    
                                '2. 기존 등록된 내용을 살펴본다.
                                '   단, 현재 자기의 필드에 있는 내용은 무시
                                    nLastSaveChk = 0
                                    nLastSaveChk = Find_Already_Save_LSN_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                    If nLastSaveChk > 0 Then
                                        sTmp = ""
                                        Select Case sWeek
                                            Case "2"
                                                sTmp = sTmp & "월"
                                            Case "3"
                                                sTmp = sTmp & "화"
                                            Case "4"
                                                sTmp = sTmp & "수"
                                            Case "5"
                                                sTmp = sTmp & "목"
                                            Case "6"
                                                sTmp = sTmp & "금"
                                            Case "7"
                                                sTmp = sTmp & "토"
                                            Case "1"
                                                sTmp = sTmp & "일"
                                        End Select
                                        sTmp = sTmp & "요일 " & sLesson & "교시에서 【같은 강의실에서 강의하는 강사】가 있습니다." & vbCrLf & "등록하시겠습니까?"
                                        
'                                        If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                            cmdSave_LSN.Enabled = True
'                                            Exit Sub
'                                        End If


                                        GoTo GONEXT     '< 다음 진행
                                        
                                    End If
                                    
                                    
                                '** 시간표 내역 등록하기 **
                                    Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                    Call Show_TMR_Tcr(sLsnCD, sWeek, sLesson)
                                
                                End If
                            '---------------------------------------------------------------------------------------
                            End If
                        Next nRecSlash
                        
                    End If
                    
                End If      '++ X값으로 들어오면 반려 ++
            '===============================================================
GONEXT:
            
            Next nCol
        Next nRow
    End With
    
    cmdSave_LSN.Enabled = True
    
End Sub


'## 전체 시간표 내역에서 보여주기
Public Sub Show_TMR_Tcr(ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRec        As Long
    Dim ni          As Long
    Dim sData       As String

    Dim nRow        As Long
    Dim nCol        As Long

    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sTmpTcrCD   As String
    Dim sTmpSubjCD  As String
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long

    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD, GET_KEAYOL_N_LSN_TCR01(A.ACID, A.LSNCD) AS DS"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & aLsnCD & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    
    DBRec.MoveFirst
    For nRec = 1 To DBRec.RecordCount Step 1
        
        If IsNull(DBRec.Fields("TCRCD")) = False And _
           IsNull(DBRec.Fields("SUBJCD")) = False And _
           IsNull(DBRec.Fields("DS")) = False Then
            
            sTcrCD = Trim(DBRec.Fields("TCRCD"))
            sSubjCD = Trim(DBRec.Fields("SUBJCD"))
            sData = Trim(DBRec.Fields("DS"))
            
            
            With sprTmr_Tcr
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow:        nChkRow = .Row
                    .Col = SpreadHeader:            sTmpTcrCD = Trim(.Text)
                    .Col = SpreadHeader + 1:        sTmpSubjCD = Trim(.Text)
                    
                    If StrComp(sTcrCD, sTmpTcrCD, vbTextCompare) = 0 And _
                       StrComp(sSubjCD, sTmpSubjCD, vbTextCompare) = 0 Then
                       
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol:        nChkCol = .Col
                            .Row = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
                            .Row = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
                            
                            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                               
                                .Row = nChkRow
                                .Col = nChkCol
                                
                                If Trim(.Text) = "" Then
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                        
                                        If InStr(1, sData, "/", vbTextCompare) > 0 Then
                                            .Row2 = .Row
                                            .Col2 = .Col
                                            .BlockMode = True
                                                .BackColor = basModule.SectionColor1
                                                .BackColorStyle = BackColorStyleUnderGrid
                                            .BlockMode = False
                                            
                                        End If
                                    End If
                                End If
                            End If
                        Next nCol
                    End If
                Next nRow
            End With
        End If
        
        DBRec.MoveNext
    Next nRec
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
        
End Sub





'<< 데이터 저장 : 과목, 강사내역으로 저장 >>
Private Sub Save_TMR_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aWeek As String, ByVal aLesson As String, ByVal aLsnCD As String, _
                          Optional aPrt_Kaeyol As String, Optional aPrt_Lsn As String, Optional aPrt_LsnNM As String)
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim nExe        As Long
    
    Dim ni          As Integer
    Dim sSaveGbn    As String
    
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD"
    sStr = sStr & "   FROM SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"
    
    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.TCRCD  = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "    AND A.SUBJCD = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & Trim(aWeek)
    sStr = sStr & "    AND A.LESSON = " & Trim(aLesson)
    sStr = sStr & "    AND A.LSNCD  = '" & Trim(aLsnCD) & "'"
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
            


    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            ' NO ACTION
            basDataBase.DBConn.RollbackTrans
            
        ElseIf .RecordCount = 0 Then
            Set DBRec = Nothing
            
            nExe = 0
            
            If aPrt_Kaeyol = "" And aPrt_Lsn = "" And aPrt_LsnNM = "" Then
                '<< INSERT >>
                sStr = ""
                sStr = sStr & "  INSERT INTO SDTRX50TB (YM, ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS ) "
                sStr = sStr & "  VALUES ("
                sStr = sStr & "         '" & Trim(fpYM.UnFmtText) & "', "
                sStr = sStr & "         '" & Trim(basModule.SchCD) & "', "
                sStr = sStr & "         '" & Trim(aTcrCD) & "', "
                sStr = sStr & "         '" & Trim(aSubjCD) & "', "
                sStr = sStr & "         '" & Trim(aLsnCD) & "', "
                sStr = sStr & "         " & Trim(aLesson) & ", "
                sStr = sStr & "         " & Trim(aWeek)
                sStr = sStr & "  ) "
            Else
                
                '<< insert : 일반항목 등록 외1 >>
                sStr = ""
                sStr = sStr & "  INSERT INTO SDTRX50TB (YM, ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS, PRT_KAEYOL, PRT_LSN, PRT_LSNNM ) "
                sStr = sStr & "  VALUES ("
                sStr = sStr & "         '" & Trim(fpYM.UnFmtText) & "', "
                sStr = sStr & "         '" & Trim(basModule.SchCD) & "', "
                sStr = sStr & "         '" & Trim(aTcrCD) & "', "
                sStr = sStr & "         '" & Trim(aSubjCD) & "', "
                sStr = sStr & "         '" & Trim(aLsnCD) & "', "
                sStr = sStr & "         " & Trim(aLesson) & ", "
                sStr = sStr & "         " & Trim(aWeek) & ", "
                
                sStr = sStr & "         '" & Trim(aPrt_Kaeyol) & "', "      ' 추가항목
                sStr = sStr & "         '" & Trim(aPrt_Lsn) & "', "
                sStr = sStr & "         '" & Trim(aPrt_LsnNM) & "' "
                
                sStr = sStr & "  ) "
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
                
                
                '<< 시간표 변경내역 처리 --------------------------------------------------------------------------
                
                Call Disp_Detail_Tmr_Data(aTcrCD, aSubjCD, aLsnCD, aWeek, aLesson)
                Call Show_TMR_Lsn(aLsnCD, aWeek, aLesson)
                
                '--------------------------------------------------------------------------------------------------
                
            End If
            
        Else
            sTmp = ""
            Select Case aWeek
                Case "2"
                    sTmp = sTmp & "월"
                Case "3"
                    sTmp = sTmp & "화"
                Case "4"
                    sTmp = sTmp & "수"
                Case "5"
                    sTmp = sTmp & "목"
                Case "6"
                    sTmp = sTmp & "금"
                Case "7"
                    sTmp = sTmp & "토"
                Case "1"
                    sTmp = sTmp & "일"
            End Select
            sTmp = sTmp & "요일 " & aLesson & "교시 등록오류입니다."
            MsgBox sTmp, vbExclamation + vbOKOnly, "시간표 등록오류"
            
            basDataBase.DBConn.RollbackTrans
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
End Sub





'## 전체 시간표 내역에서 보여주기
Public Sub Show_TMR_Lsn(ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRec        As Long
    Dim ni          As Long
    Dim sData       As String

    Dim nRow        As Long
    Dim nCol        As Long

    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long

    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD)||','||GET_TCRNM(A.ACID, A.TCRCD) AS DS"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"
    
    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & aLsnCD & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        .MoveFirst
        
        sData = ""
        For nRec = 1 To .RecordCount Step 1
            If IsNull(.Fields("DS")) = False Then
                If sData <> "" Then
                    sData = sData & "/"
                End If
                sData = sData & Trim(.Fields("DS"))
            End If
            
            .MoveNext
        Next nRec
    End With
    
    
    With sprTmr_Lsn
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow:        nChkRow = .Row
            .Col = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
            .Col = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
            
            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                    For nCol = 1 To .MaxCols Step 1
                        .Col = nCol:        nChkCol = .Col
                        .Row = SpreadHeader + 1
                        
                        If StrComp(aLsnCD, Trim(.Text), vbTextCompare) = 0 Then
                            
                            .Row = nChkRow
                            .Col = nChkCol
                            
                                Call basFunction.Set_SprType_Text(sprTmr_Lsn, "center", "left", 60, sData)
                                
                            If InStr(1, sData, "/", vbTextCompare) > 0 Then
                                .Row2 = .Row
                                .Col2 = .Col
                                .BlockMode = True
                                    .BackColor = basModule.SectionColor1
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                            End If
                            
                        End If
                    Next nCol
            End If
        Next nRow
    End With

ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
        
End Sub



































'## 저장데이터 보여주기
Private Sub Disp_Detail_Tmr_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String)
      
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim ni          As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    
    On Error GoTo ErrStmt
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, GET_TCRNM(A.ACID, A.TCRCD) AS TCRNM,"
    sStr = sStr & "        A.SUBJCD, GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD) AS SUBJNM,"
    sStr = sStr & "        GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD)||','||GET_TCRNM(A.ACID, A.TCRCD) AS LSNDATA,"
    sStr = sStr & "        GET_KEAYOL_N_LSN_TCR01(A.ACID, A.LSNCD) AS LSNCDNM,"
    sStr = sStr & "        A.LESSON, A.WEEKS"
    sStr = sStr & "   FROM SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.TCRCD  = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "    AND A.SUBJCD = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & Trim(aWeek)
    sStr = sStr & "    AND A.LESSON = " & Trim(aLesson)
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
            
 
 
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            
            ' 요일,교시 & 반 내역 등록
            For nRow = 1 To sprTmr_Lsn.MaxRows Step 1
                sprTmr_Lsn.Row = nRow
                sprTmr_Lsn.Col = SpreadHeader + 1        '< 요일
                
                If StrComp(Trim(sprTmr_Lsn.Text), aWeek, vbTextCompare) = 0 Then
                    nr_Chk = sprTmr_Lsn.Row              '< row 값

                    sprTmr_Lsn.Col = SpreadHeader + 2        '< lesson
                    
                    If StrComp(Trim(sprTmr_Lsn.Text), aLesson, vbTextCompare) = 0 Then
                        
                        For nCol = 1 To sprTmr_Lsn.MaxCols Step 1
                            sprTmr_Lsn.Col = nCol
                            sprTmr_Lsn.Row = SpreadHeader + 1
                        
                            If StrComp(Trim(sprTmr_Lsn.Text), aLsnCD, vbTextCompare) = 0 Then
                                nc_Chk = sprTmr_Lsn.Col
                                
                                sprTmr_Lsn.Row = nr_Chk
                                sprTmr_Lsn.Col = nc_Chk
                                If sprTmr_Lsn.Text = "" Then
                                    sprTmr_Lsn.Text = Trim(.Fields("LSNDATA"))
                                Else
                                    sprTmr_Lsn.Text = Trim(.Fields("LSNDATA")) & "/" & Trim(sprTmr_Lsn.Text)
                                End If
                                    
                                Exit For
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
            ' 강사 & 요일 내역 등록
            For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
                sprTmr_Tcr.Row = nRow
                sprTmr_Tcr.Col = SpreadHeader
                
                If StrComp(Trim(sprTmr_Tcr.Text), aTcrCD, vbTextCompare) = 0 Then
                    sprTmr_Tcr.Col = SpreadHeader + 1
                    
                    If StrComp(Trim(sprTmr_Tcr.Text), aSubjCD, vbTextCompare) = 0 Then
                        nr_Chk = sprTmr_Tcr.Row
                        
                        sprTmr_Tcr.Col = SpreadHeader + 6           '< 변경내역
                            sprTmr_Tcr.Text = Trim(CStr(CLng(sprTmr_Tcr.Text) - 1))
                        
                        sprTmr_Tcr.Row = nr_Chk
                        For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
                            sprTmr_Tcr.Col = nCol
                            sprTmr_Tcr.Row = SpreadHeader + 1
                            
                            If StrComp(Trim(sprTmr_Tcr.Text), aWeek, vbTextCompare) = 0 Then
                                sprTmr_Tcr.Row = SpreadHeader + 2
                                
                                If StrComp(Trim(sprTmr_Tcr.Text), aLesson, vbTextCompare) = 0 Then
                                    nc_Chk = sprTmr_Tcr.Col
                                    
                                    sprTmr_Tcr.Row = nr_Chk
                                    sprTmr_Tcr.Col = nc_Chk
                                    If sprTmr_Tcr.Text = "" Then
                                        sprTmr_Tcr.Text = Trim(.Fields("LSNCDNM"))
                                    Else
                                        sprTmr_Tcr.Text = Trim(.Fields("LSNCDNM")) & "/" & Trim(sprTmr_Tcr.Text)
                                    End If
                                        
                                    Exit For
                                End If
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
            
            ' 강사 시수내역 조정하기
            sprGwamok.MaxRows = 0                           '< 강사, 과목내역 초기화
            For nRow = 1 To sprSisu.MaxRows Step 1          '< 요일별 내역 초기화
                For nCol = 1 To sprSisu.MaxCols Step 1
                    sprSisu.Row = nRow
                    sprSisu.Col = nCol
                        sprSisu.Text = ""
                Next nCol
            Next nRow
            
            sprSisu.Row = 1:        sprSisu.Row2 = sprSisu.MaxRows
            sprSisu.Col = 1:        sprSisu.Col2 = sprSisu.MaxCols
            sprSisu.BlockMode = True
                sprSisu.BackColor = basModule.WhiteColor
                sprSisu.BackColorStyle = BackColorStyleUnderGrid
            sprSisu.BlockMode = False
                            
            For nRow = 1 To sprTcr.MaxRows Step 1           '< 강사시수처리
                sprTcr.Row = nRow
                sprTcr.Col = 2
                If StrComp(aTcrCD, Trim(sprTcr.Text), vbTextCompare) = 0 Then
                    Select Case aWeek
                        Case "2"
                            sprTcr.Col = 6
                            fpT(2).Value = fpT(2).Value + 1
                        Case "3"
                            sprTcr.Col = 7
                            fpT(3).Value = fpT(3).Value + 1
                        Case "4"
                            sprTcr.Col = 8
                            fpT(4).Value = fpT(4).Value + 1
                        Case "5"
                            sprTcr.Col = 9
                            fpT(5).Value = fpT(5).Value + 1
                        Case "6"
                            sprTcr.Col = 10
                            fpT(6).Value = fpT(6).Value + 1
                        Case "7"
                            sprTcr.Col = 11
                            fpT(7).Value = fpT(7).Value + 1
                        Case "1"
                            sprTcr.Col = 12
                    End Select
                    
                    If Trim(sprTcr.Text) = "" Then
                        Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", 1)
                    Else
                        Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) + 1)
                    End If
                    
                    sprTcr.Col = 5
                        fpT(1).Value = fpT(1).Value - 1
                    Call basFunction.Set_SprType_Numeric(sprTcr, 0, -9999, 9999, "", CLng(sprTcr.Text) - 1)
                    
                End If
            Next nRow
            
        End If
    End With
    
    'MsgBox "저장하였습니다.", vbInformation + vbOKOnly, "시간표 저장하기"
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
End Sub


















'<< 이미 저장된 데이터가 있는지 찾음 : 다른 강사가 인지 체크 >>
Private Function Find_Already_Save_LSN_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aWeek As String, ByVal aLesson As String, ByVal aLsnCD As String) As Long

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRet        As Long
    Dim ni          As Long

    On Error GoTo ErrStmt
    
    nRet = 0

    sStr = ""
    sStr = sStr & " SELECT COUNT(*) AS CNT"
    sStr = sStr & "   FROM ("
    sStr = sStr & "         SELECT A.*"
    sStr = sStr & "           From SDTRX50TB A,"
    
    sStr = sStr & "                (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                   FROM SDLSN01TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                 UNION"
    sStr = sStr & "                 SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                   FROM SDLSN02TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                ) B"
    
    sStr = sStr & "          WHERE A.ACID   = B.ACID  "
    sStr = sStr & "            AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "            AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "            AND A.ACID     = '" & Trim(basModule.SchCD) & "'"
    '                           // TCRCD //
    sStr = sStr & "            AND A.SUBJCD   = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "            AND A.WEEKS    = " & Trim(aWeek)
    sStr = sStr & "            AND A.LESSON   = " & Trim(aLesson)
    sStr = sStr & "            AND A.LSNCD    = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "         MINUS"
    sStr = sStr & "         SELECT A.*"
    sStr = sStr & "           From SDTRX50TB A, "
    sStr = sStr & "                (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                   FROM SDLSN01TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                 UNION"
    sStr = sStr & "                 SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                   FROM SDLSN02TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                ) B"
    sStr = sStr & "          WHERE A.ACID     = B.ACID  "
    sStr = sStr & "            AND A.LSNCD    = B.LSNCD "
    sStr = sStr & "            AND A.YM       = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "            AND A.ACID     = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND A.TCRCD    = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "            AND A.SUBJCD   = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "            AND A.WEEKS    = " & Trim(aWeek)
    sStr = sStr & "            AND A.LESSON   = " & Trim(aLesson)
    sStr = sStr & "            AND A.LSNCD    = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "         )"
    
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        .MoveFirst
        
        If .RecordCount = 1 Then
            
            If IsNull(.Fields("CNT")) = False Then
                If IsNumeric(.Fields("CNT")) = True Then
                    nRet = CLng(.Fields("CNT"))
                End If
            End If
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Find_Already_Save_LSN_Data = nRet

    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    On Error GoTo 0
    
    Find_Already_Save_LSN_Data = 0
    
End Function



'<< 이미 저장된 데이터가 있는지 찾음 : 다른반에서 강의를 하는 강사인지 체크 >>
Private Function Find_Already_Save_TCR_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aWeek As String, ByVal aLesson As String, ByVal aLsnCD As String) As Long

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRet        As Long
    Dim ni          As Long

    On Error GoTo ErrStmt
    
    nRet = 0

    sStr = ""
    sStr = sStr & " SELECT COUNT(*) AS CNT"
    sStr = sStr & "   FROM ("
    sStr = sStr & "         SELECT A.*"
    sStr = sStr & "           From SDTRX50TB A, "
    sStr = sStr & "                (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                   FROM SDLSN01TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                 UNION"
    sStr = sStr & "                 SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                   FROM SDLSN02TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                ) B"
    sStr = sStr & "          WHERE A.ACID     = B.ACID  "
    sStr = sStr & "            AND A.LSNCD    = B.LSNCD "
    sStr = sStr & "            AND A.YM       = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "            AND A.ACID     = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND A.TCRCD    = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "            AND A.SUBJCD   = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "            AND A.WEEKS    = " & Trim(aWeek)
    sStr = sStr & "            AND A.LESSON   = " & Trim(aLesson)
    '                           // LSNCD //
    sStr = sStr & "         MINUS"
    sStr = sStr & "         SELECT A.*"
    sStr = sStr & "           From SDTRX50TB A,"
    sStr = sStr & "                (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "                   FROM SDLSN01TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                 UNION"
    sStr = sStr & "                 SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                   FROM SDLSN02TB "
    sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                ) B"
    sStr = sStr & "          WHERE A.ACID     = B.ACID  "
    sStr = sStr & "            AND A.LSNCD    = B.LSNCD "
    sStr = sStr & "            AND A.YM       = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "            AND A.ACID     = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND A.TCRCD    = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "            AND A.SUBJCD   = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "            AND A.WEEKS    = " & Trim(aWeek)
    sStr = sStr & "            AND A.LESSON   = " & Trim(aLesson)
    sStr = sStr & "            AND A.LSNCD    <> '" & Trim(aLsnCD) & "'"
    sStr = sStr & "         )"
    
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        .MoveFirst
        
        If .RecordCount = 1 Then
            
            If IsNull(.Fields("CNT")) = False Then
                If IsNumeric(.Fields("CNT")) = True Then
                    nRet = CLng(.Fields("CNT"))
                End If
            End If
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Find_Already_Save_TCR_Data = nRet

    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    On Error GoTo 0
    
    Find_Already_Save_TCR_Data = 0
    
End Function


'<< 과목 및 강사 조회 >>
Private Sub Find_Tcr_and_Subj_Code(ByRef aTcrCD As String, ByRef aSubjCD As String, ByVal aTcrNM As String, ByVal aSubjNM As String)

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String
    
    Dim ni          As Long

    On Error GoTo ErrStmt

    sStr = ""
    sStr = sStr & "    SELECT TCRCD, SUBJCD"
    sStr = sStr & "      From SDTCR01TB"
    sStr = sStr & "     WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "       AND TRIM(TCRNM)  = '" & aTcrNM & "'"
    sStr = sStr & "       AND TRIM(SUBJNM) = '" & aSubjNM & "'"
    sStr = sStr & "       AND ROWNUM = 1 "

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        .MoveFirst
        
        If .RecordCount = 1 Then
            aTcrCD = "":    If IsNull(.Fields("TCRCD")) = False Then aTcrCD = Trim(.Fields("TCRCD"))
            aSubjCD = "":   If IsNull(.Fields("SUBJCD")) = False Then aSubjCD = Trim(.Fields("SUBJCD"))
            
            .MoveNext
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    On Error GoTo 0
    MsgBox "강사 및 과목 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목조회"

End Sub









'############################################### 등록 작업 : 강사별 내역으로 저장 ################################################################################


'## 각 데이터 저장내역을 살펴보고, 바뀐내역 또는 없는 내역은 신규 또는 갱신함.
'## 시간은 걸리겠지만, 등록하는 시점에서 저장하는 것이 해당 프로그램에선 불가능


Private Sub cmdSave_Tcr_Click()
    Dim nRow            As Long
    Dim nCol            As Long
    
    Dim sTmp            As String
    
    Dim sDivSlash()     As String
    Dim nRecSlash       As Long
    
    Dim sDivComma()     As String
    Dim nRecComma       As Long
    
    Dim nChkRow         As Long
    Dim nChkCol         As Long
    
    Dim sTmpLsn         As String
    Dim sTmpKaeyol      As String
    
    Dim sLsnCD          As String
    Dim sWeek           As String
    Dim sLesson         As String
    
    Dim sTcrCD          As String
    Dim sSubjCD         As String
    
    Dim sPrt_Kaeyol     As String
    Dim sPrt_Lsn        As String
    Dim sPrt_LsnNM      As String
    
    Dim nLastSaveChk    As Long
    
    
    cmdSave_Tcr.Enabled = False
    
    
    With sprTmr_Tcr
        
        ProgressBar1.Min = 0
        ProgressBar1.Max = 100
        ProgressBar1.Value = 0
        
        For nRow = 1 To .MaxRows Step 1
            
            ProgressBar1.Value = Fix(nRow / .MaxRows * 100)
            
            For nCol = 1 To .MaxCols Step 1
            
                .Row = nRow:        nChkRow = .Row
                .Col = nCol:        nChkCol = .Col
                
                If StrComp(Trim(.Text), "X", vbTextCompare) <> 0 Then       ' X값으로 들어오면 반려
                
                '===============================================================
                '## 일반적 데이터
                '===============================================================
                    If InStr(1, Trim(.Text), "/", vbTextCompare) = 0 Then
                        
                        If Len(Trim(.Text)) >= 3 Then               ' 3자리이상 : 1(계열) 2,3 (반 코드명칭)
                                                                    '   추가 - 501, 계열, 표시반명
                            
                            sDivComma() = Split(UCase(Trim(.Text)), ",", -1, vbTextCompare)
                            If UBound(sDivComma) < 1 Then   '<< 일반적인 경우 : , 가 없음 - 그냥 3자리가 있음.
                                
                                sTmpLsn = Right(Left(Trim(.Text), 3), 2)
                                sTmpKaeyol = "0" & Left(Trim(.Text), 1)
                                
                                Select Case sTmpKaeyol
                                    Case "01"
                                        Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)
                                        
                                        sPrt_Kaeyol = ""
                                        sPrt_Lsn = ""
                                        sPrt_LsnNM = ""
                                    Case "02"
                                        Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)
                                        
                                        sPrt_Kaeyol = ""
                                        sPrt_Lsn = ""
                                        sPrt_LsnNM = ""
                                        
'                                    Case "05"
'
'                                        sLsnCD = "00000"
'
'                                        sPrt_Kaeyol = "1"
'                                        sPrt_Lsn = Left(Trim(.Text), 3)
'                                        sPrt_LsnNM = "임시1"
'                                    Case "06"
'
'                                        sLsnCD = "00000"
'
'                                        sPrt_Kaeyol = "2"
'                                        sPrt_Lsn = Left(Trim(.Text), 3)
'                                        sPrt_LsnNM = "임시2"

                                    Case Else
                                    
                                        sLsnCD = "00000"
                                        
                                        sPrt_Kaeyol = "1"
                                        sPrt_Lsn = Left(Trim(.Text), 3)
                                        sPrt_LsnNM = "임시ext"
                                End Select
                            Else            '<< 【 501, 계열, 표시반명 】인 경우
                            
'                                sLsnCD = "00000"
'                                sPrt_Kaeyol = Trim(sDivComma(1))
'                                sPrt_Lsn = Trim(sDivComma(0))
'                                sPrt_LsnNM = Trim(sDivComma(2))
                                
                            End If
                            
                            '   SLSNCD
                            '   sPrt_Kaeyol
                            '   sPrt_Lsn
                            '   sPrt_LsnNM
                                        
                            .Row = nChkRow
                                .Col = SpreadHeader:            sTcrCD = Trim(.Text)
                                .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
                            .Col = nChkCol
                                .Row = SpreadHeader + 1:        sWeek = Trim(.Text)
                                .Row = SpreadHeader + 2:        sLesson = Trim(.Text)
                            
                            
                        '> 조회된 반 내역이 있어야 함. ------------------------------------------------
                            If sLsnCD <> "" Then
                            
                            '1. 기존 등록된 내용을 살펴본다.
                            '   단, 현재 자기의 필드에 있는 내용은 무시
                                nLastSaveChk = 0
                                nLastSaveChk = Find_Already_Save_TCR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)      '< 공통함수 사용
                                If nLastSaveChk > 0 Then
                                    sTmp = ""
                                    Select Case sWeek
                                        Case "2"
                                            sTmp = sTmp & "월"
                                        Case "3"
                                            sTmp = sTmp & "화"
                                        Case "4"
                                            sTmp = sTmp & "수"
                                        Case "5"
                                            sTmp = sTmp & "목"
                                        Case "6"
                                            sTmp = sTmp & "금"
                                        Case "7"
                                            sTmp = sTmp & "토"
                                        Case "1"
                                            sTmp = sTmp & "일"
                                    End Select
                                    sTmp = sTmp & "요일 " & sLesson & "교시에서 【다른반에 강의】를 합니다." & vbCrLf & "등록하시겠습니까?"
                                    
'                                    If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                        cmdSave_Tcr.Enabled = True
'                                        Exit Sub
'                                    End If
                                    
                                    GoTo GONEXT     '< 다음 진행
                                    
                                End If
                                
                            '2. 기존 등록된 내용을 살펴본다.
                            '   단, 현재 자기의 필드에 있는 내용은 무시
                                nLastSaveChk = 0
                                nLastSaveChk = Find_Already_Save_LSN_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)      '< 공통함수 사용
                                If nLastSaveChk > 0 Then
                                    sTmp = ""
                                    Select Case sWeek
                                        Case "2"
                                            sTmp = sTmp & "월"
                                        Case "3"
                                            sTmp = sTmp & "화"
                                        Case "4"
                                            sTmp = sTmp & "수"
                                        Case "5"
                                            sTmp = sTmp & "목"
                                        Case "6"
                                            sTmp = sTmp & "금"
                                        Case "7"
                                            sTmp = sTmp & "토"
                                        Case "1"
                                            sTmp = sTmp & "일"
                                    End Select
                                    sTmp = sTmp & "요일 " & sLesson & "교시에서 【같은 강의실에서 강의하는 강사】가 있습니다." & vbCrLf & "등록하시겠습니까?"
                                    
'                                    If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                        cmdSave_Tcr.Enabled = True
'                                        Exit Sub
'                                    End If

                                    GoTo GONEXT     '< 다음 진행
                                    
                                End If
                                
                                
                            '** 시간표 내역 등록하기 **
                                Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD, sPrt_Kaeyol, sPrt_Lsn, sPrt_LsnNM)
                                
                            
                            End If
                        End If
                        
                '===============================================================
                '## 일반적 데이터
                '===============================================================
                    Else
                        
                        sDivSlash() = Split(Trim(.Text), "/", -1, vbTextCompare)
                        For nRecSlash = 0 To UBound(sDivSlash) Step 1

                            If Len(sDivSlash(nRecSlash)) >= 3 Then               ' 3자리이상 : 1(계열) 2,3 (반 코드명칭)
                                                                        '   추가 - 501, 계열, 표시반명
                                                                    
                                sDivComma() = Split(Trim(sDivSlash(nRecSlash)), ",", -1, vbTextCompare)

                                If UBound(sDivComma) < 1 Then   '<< 일반적인 경우 : , 가 없음

                                    sTmpLsn = Right(Left(Trim(sDivSlash(nRecSlash)), 3), 2)
                                    sTmpKaeyol = "0" & Left(Trim(sDivSlash(nRecSlash)), 1)

                                    Select Case sTmpKaeyol
                                        Case "01"
                                            Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)

                                            sPrt_Kaeyol = ""
                                            sPrt_Lsn = ""
                                            sPrt_LsnNM = ""
                                        Case "02"
                                            Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)

                                            sPrt_Kaeyol = ""
                                            sPrt_Lsn = ""
                                            sPrt_LsnNM = ""
                                            
                                            
'                                        Case "05"
'
'                                            sLsnCD = "00000"
'
'                                            sPrt_Kaeyol = "1"
'                                            sPrt_Lsn = Left(Trim(sDivSlash(nRecSlash)), 3)
'                                            sPrt_LsnNM = "임시1"
'                                        Case "06"
'
'                                            sLsnCD = "00000"
'
'                                            sPrt_Kaeyol = "2"
'                                            sPrt_Lsn = Left(Trim(sDivSlash(nRecSlash)), 3)
'                                            sPrt_LsnNM = "임시2"
                                            
                                        
                                        Case Else
                                    
                                            sLsnCD = "00000"
                                            
                                            sPrt_Kaeyol = "1"
                                            sPrt_Lsn = Left(Trim(sDivSlash(nRecSlash)), 3)
                                            sPrt_LsnNM = "임시ext"
                                        
                                    End Select
                                Else            '<< 【 501, 계열, 표시반명 】인 경우

'                                    sLsnCD = "00000"
'                                    sPrt_Kaeyol = Trim(sDivComma(1))
'                                    sPrt_Lsn = Trim(sDivComma(0))
'                                    sPrt_LsnNM = Trim(sDivComma(2))

                                End If

                                    '   SLSNCD
                                    '   sPrt_Kaeyol
                                    '   sPrt_Lsn
                                    '   sPrt_LsnNM

                                .Row = nChkRow
                                    .Col = SpreadHeader:            sTcrCD = Trim(.Text)
                                    .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
                                .Col = nChkCol
                                    .Row = SpreadHeader + 1:        sWeek = Trim(.Text)
                                    .Row = SpreadHeader + 2:        sLesson = Trim(.Text)


                                '> 조회된 반 내역이 있어야 함. ------------------------------------------------
                                    If sLsnCD <> "" Then

                                    '1. 기존 등록된 내용을 살펴본다.
                                    '   단, 현재 자기의 필드에 있는 내용은 무시
                                        nLastSaveChk = 0
                                        nLastSaveChk = Find_Already_Save_TCR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)      '< 공통함수 사용
                                        If nLastSaveChk > 0 Then
                                            sTmp = ""
                                            Select Case sWeek
                                                Case "2"
                                                    sTmp = sTmp & "월"
                                                Case "3"
                                                    sTmp = sTmp & "화"
                                                Case "4"
                                                    sTmp = sTmp & "수"
                                                Case "5"
                                                    sTmp = sTmp & "목"
                                                Case "6"
                                                    sTmp = sTmp & "금"
                                                Case "7"
                                                    sTmp = sTmp & "토"
                                                Case "1"
                                                    sTmp = sTmp & "일"
                                            End Select
                                            sTmp = sTmp & "요일 " & sLesson & "교시에서 【다른반에 강의】를 합니다." & vbCrLf & "등록하시겠습니까?"

'                                            If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                                cmdSave_Tcr.Enabled = True
'                                                Exit Sub
'                                            End If

                                            GoTo GONEXT     '< 다음 진행
                                            
                                        End If

                                    '2. 기존 등록된 내용을 살펴본다.
                                    '   단, 현재 자기의 필드에 있는 내용은 무시
                                        nLastSaveChk = 0
                                        nLastSaveChk = Find_Already_Save_LSN_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)      '< 공통함수 사용
                                        If nLastSaveChk > 0 Then
                                            sTmp = ""
                                            Select Case sWeek
                                                Case "2"
                                                    sTmp = sTmp & "월"
                                                Case "3"
                                                    sTmp = sTmp & "화"
                                                Case "4"
                                                    sTmp = sTmp & "수"
                                                Case "5"
                                                    sTmp = sTmp & "목"
                                                Case "6"
                                                    sTmp = sTmp & "금"
                                                Case "7"
                                                    sTmp = sTmp & "토"
                                                Case "1"
                                                    sTmp = sTmp & "일"
                                            End Select
                                            sTmp = sTmp & "요일 " & sLesson & "교시에서 【같은 강의실에서 강의하는 강사】가 있습니다." & vbCrLf & "등록하시겠습니까?"

'                                            If MsgBox(sTmp, vbQuestion + vbYesNo, "시간표 등록") = vbNo Then
'                                                cmdSave_Tcr.Enabled = True
'                                                Exit Sub
'                                            End If

                                            GoTo GONEXT     '< 다음 진행
                                            
                                        End If


                                    '** 시간표 내역 등록하기 **
                                        Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD, sPrt_Kaeyol, sPrt_Lsn, sPrt_LsnNM)


                                    End If

                            End If
                        Next nRecSlash
                        
                    End If
                    
                End If      '++ X값으로 들어오면 반려 ++
            '===============================================================
GONEXT:
            
            Next nCol
        Next nRow
    End With
    
    cmdSave_Tcr.Enabled = True
    
End Sub


Private Sub Get_LsnCD_Data(ByRef aLsnCD As String, ByVal aKaeyol As String, ByVal aLsn As String)

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String
    
    Dim ni          As Long

    On Error GoTo ErrStmt

    sStr = ""
    sStr = sStr & " SELECT LSNCD"
    sStr = sStr & "   FROM (SELECT LSNCD"                                       '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB"
    sStr = sStr & "          WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND KAEYOL  = '" & aKaeyol & "'"
    sStr = sStr & "            AND LSNCDNM = '" & aLsn & "'"
    sStr = sStr & "         UNION ALL"
    sStr = sStr & "         SELECT LSNCD"
    sStr = sStr & "           FROM SDLSN02TB"
    sStr = sStr & "          WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND KAEYOL  = '" & aKaeyol & "'"
    sStr = sStr & "            AND LSNCDNM = '" & aLsn & "'"
    sStr = sStr & "        )"
    sStr = sStr & "  GROUP BY LSNCD"

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        .MoveFirst
        
        If .RecordCount = 1 Then
            aLsnCD = "":    If IsNull(.Fields("LSNCD")) = False Then aLsnCD = Trim(.Fields("LSNCD"))
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

    On Error GoTo 0
    MsgBox "강사 및 과목 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목조회"

End Sub

























'------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 시간표 배분작업 ################################################################################

'## 시간표 자동배분하기
Private Sub cmdAutoTmr_Click()
    Dim nRow        As Long
    
    fraAuto.Visible = True
    fraAuto.ZOrder 0
    
    ProgressBar2.Min = 0
    ProgressBar2.Max = 100
    ProgressBar2.Value = 0
    
    With sprAutoGwamokSort
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = 1
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    'Call cmdCalcu_TCR_Click         '< 강사현황
    
End Sub


'## 강사별 시수내역 조회
Private Sub cmdCalcu_TCR_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "        SELECT ACID, TCRCD, SUBJCD, LSNCD, SUBJGBN, "
    sStr = sStr & "               GET_TCRNM(ACID, TCRCD) AS TCRNM, GET_SUBJNM(ACID, TCRCD, SUBJCD) AS SUBJNM, "
    sStr = sStr & "               GET_LSNNM(ACID, LSNCD) AS LSNNM, "
    sStr = sStr & "               SISU"
    sStr = sStr & "          FROM (SELECT ACID, TCRCD, SUBJCD, MAX(SUBJGBN) AS SUBJGBN, LSNCD, SUM(T_SISU)-SUM(S_SISU) AS SISU, GET_KEAYOL_N_LSN_TCR01(ACID, LSNCD) AS KAEYOL, MAX(SUBJORD) AS SUBJORD"
    sStr = sStr & "                  FROM ("
    sStr = sStr & "                        SELECT A.ACID, A.TCRCD, A.SUBJCD, MAX(A.SUBJGBN) AS SUBJGBN, B.LSNCD, SUM(B.SISU) AS T_SISU, 0 AS S_SISU,"
    '>> 순서
    With sprAutoGwamokSort
        .Row = 1:   .Col = 2:   sStr = sStr & "   DECODE(MAX(A.SUBJGBN),10," & Trim(CStr(.Text))
        .Row = 2:   .Col = 2:   sStr = sStr & "                        ,20," & Trim(CStr(.Text))
        .Row = 3:   .Col = 2:   sStr = sStr & "                        ,30," & Trim(CStr(.Text))
        .Row = 4:   .Col = 2:   sStr = sStr & "                        ,40," & Trim(CStr(.Text))
        .Row = 5:   .Col = 2:   sStr = sStr & "                        ,50," & Trim(CStr(.Text))
                                sStr = sStr & "   ) AS SUBJORD "
    End With
    sStr = sStr & "                          FROM SDTCR01TB A, SDTCR11TB B, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) C"
    
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    
    sStr = sStr & "                           AND B.ACID   = C.ACID  "
    sStr = sStr & "                           AND B.LSNCD  = C.LSNCD "
    
    sStr = sStr & "                           AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         GROUP BY A.ACID, A.TCRCD, A.SUBJCD, B.LSNCD"
    sStr = sStr & "                        UNION ALL"
    sStr = sStr & "                        SELECT A.ACID, A.TCRCD, A.SUBJCD, '' AS SUBJGBN, A.LSNCD, 0 AS T_SISU, SUM(A.SISU) AS S_SISU, 0 AS SUBJORD "
    sStr = sStr & "                          FROM SDTRX50TB A, "
    
    sStr = sStr & "                               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
    sStr = sStr & "                                  FROM SDLSN01TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                UNION"
    sStr = sStr & "                                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "                                  FROM SDLSN02TB "
    sStr = sStr & "                                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                               ) B"
    
    sStr = sStr & "                         WHERE A.ACID = B.ACID  "
    sStr = sStr & "                           AND A.LSNCD= B.LSNCD "
    sStr = sStr & "                           AND A.YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         GROUP BY A.ACID, A.TCRCD, A.SUBJCD, A.LSNCD"
    sStr = sStr & "                        )"
    sStr = sStr & "                 GROUP BY ACID, TCRCD, SUBJCD, LSNCD"
    sStr = sStr & "               )"
    sStr = sStr & "         WHERE SISU > 0"
    If Trim(Right(cboAutoTmrGbn.Text, 30)) <> "ALL" Then
        Select Case Trim(Right(cboAutoTmrGbn.Text, 30))
            Case "TAM"
                sStr = sStr & " AND TCRCD "
                sStr = sStr & "  IN ("
                sStr = sStr & "      SELECT TCRCD "
                sStr = sStr & "        From SDTCR01TB"
                sStr = sStr & "       WHERE ACID = '" & Trim(basModule.SchCD) & "'"
                sStr = sStr & "         AND SUBJGBN IN ('40','50')"
                sStr = sStr & "      )"
            Case "KME"
                sStr = sStr & " AND TCRCD "
                sStr = sStr & "  IN ("
                sStr = sStr & "      SELECT TCRCD "
                sStr = sStr & "        From SDTCR01TB"
                sStr = sStr & "       WHERE ACID = '" & Trim(basModule.SchCD) & "'"
                sStr = sStr & "         AND SUBJGBN IN ('10','20','30')"
                sStr = sStr & "      )"
        End Select
    End If
    sStr = sStr & "         ORDER BY ACID, SUBJORD, SUBJGBN, TCRCD, KAEYOL, SUBJCD, LSNCD"
    
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
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sprAutoTeacher.MaxRows = 0
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            ProgressBar2.Value = 0
            If .RecordCount = 0 Then
                ProgressBar1.Value = 100
            End If
            
            For nRec = 1 To .RecordCount Step 1
            
                ProgressBar2.Value = Fix(nRec / .RecordCount * 100)
            
                sprAutoTeacher.MaxRows = sprAutoTeacher.MaxRows + 1
                sprAutoTeacher.Row = sprAutoTeacher.MaxRows
                
                sprAutoTeacher.Col = 1
                    sTmp = " ":     If IsNull(.Fields("ACID")) = False Then sTmp = Trim(.Fields("ACID"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprAutoTeacher.SetCellBorder sprAutoTeacher.Col, 1, sprAutoTeacher.Col, sprAutoTeacher.MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SUBJGBN")) = False Then sTmp = Trim(.Fields("SUBJGBN"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprAutoTeacher.SetCellBorder sprAutoTeacher.Col, 1, sprAutoTeacher.Col, sprAutoTeacher.MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    sTmp = " ":     If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprAutoTeacher, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                
                sprAutoTeacher.SetCellBorder sprAutoTeacher.Col, 1, sprAutoTeacher.Col, sprAutoTeacher.MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    nTmp = 0:       If IsNumeric(.Fields("SISU")) = True Then nTmp = CLng(.Fields("SISU"))          '< 작업해야 할 시수
                        Call basFunction.Set_SprType_Numeric(sprAutoTeacher, 0, 0, 99999, "", nTmp)
                
                sprAutoTeacher.Col = sprAutoTeacher.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprAutoTeacher)
                    sprAutoTeacher.Value = 0
                
                
                .MoveNext
            Next nRec
        End If
    End With
    
    With sprAutoTeacher
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "강사별 과목내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사별 과목내역"

End Sub






'==========================================================================================================================================================
'## << 배정작업 시작 >> ##
'==========================================================================================================================================================
'>> 선택 ## multi 선택
Private Sub sprAutoTeacher_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    
    If Row < 1 Then Exit Sub

    With sprAutoTeacher
        If .MaxRows < 1 Then Exit Sub

        sprAutoTeacher.Enabled = False
        
            If .Tag = "0" Then
                .Row = CLng(.Tag):      .Row2 = .Row
                .Col = 1:               .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = Row
                    .Col = .MaxCols
                    .Value = 0
                
'                For nRow = 1 To .MaxRows Step 1
'                    .Row = nRow
'                    .Col = .MaxCols
'                        .Value = 0
'                Next nRow
                
                .Row = Row:     .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = .MaxCols:    .Value = 1
                
                .Tag = Trim(CStr(Row))
            ElseIf .Tag > "0" Then
                .Row = Row
                .Col = .MaxCols
                If .Value = 1 Then
                    .Value = 0
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                Else
                    .Value = 1
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                End If
            
            End If
            
        sprAutoTeacher.Enabled = True

        sprAutoTeacher.SetFocus
        'sprAutoTeacher.SetActiveCell Col, Row

    End With
End Sub

Private Sub sprAutoTeacher_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nS      As Long
    Dim nE      As Long
    
    Dim nRow    As Long
    
    With sprAutoTeacher
    
        If .MaxRows = 0 Then Exit Sub
        
        Select Case Shift
'            Case 0
'                Call sprAutoTeacher_Click(1, .ActiveRow)
                
            Case 1          '<< shift
                If Button = vbLeftButton Then
                    If .Tag > "0" Then              '<< 1. 선택하고 2. shift를 눌러 멀티로 선택한 경우
                        nS = CLng(.Tag)
                        nE = .ActiveRow
                        
                        If nS > nE Then
                            nS = .ActiveRow
                            nE = CLng(.Tag)
                        End If
                        
                        .Row = nS:  .Row2 = nE
                        .Col = 1:   .Col2 = .MaxCols
                        .BlockMode = True
                            .BackColor = basModule.SelectColor2
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        For nRow = nS To nE Step 1
                            .Row = nRow
                            .Col = .MaxCols
                                .Value = 1
                        Next nRow
                        
                        .Tag = "0"
                        
                    End If
                End If
            
        End Select
    
    End With
End Sub

'>> 전체선택
Private Sub chkAll_Click()
    Dim ni      As Long
    
    With sprAutoTeacher
        If .MaxRows = 0 Then Exit Sub
            
        If chkAll.Value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .Value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .Value = 1
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
        
    End With
End Sub





'###########################################################
'## 강사 시간표 자동배분하기
'###########################################################
Private Sub cmdWork_Click()
    
    Dim nRow        As Long
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    Dim sLsnCD      As String
    
    Dim sTcrNM      As String
    Dim sSubjNM     As String
    Dim sTmrTcrSubj As String       '< "과목,강사"
    
    Dim nSisu       As Long
    
    Dim nRowj       As Long
    Dim nCol        As Long
    
    If MsgBox("자동배분할 강사를 선택하셨습니까?" & vbCrLf & _
              "선택하였으면 확인을 클릭하십시요." & vbCrLf & _
              "과목별 순서를 정하세요.", vbQuestion + vbYesNo, "강사별 시수현황") = vbNo Then
        Exit Sub
    End If
    
    ProgressBar2.Value = 0
    
    For nRow = 1 To sprAutoTeacher.MaxRows Step 1
    
        ProgressBar2.Value = Fix(nRow / sprAutoTeacher.MaxRows * 100)
    
        sprAutoTeacher.Row = nRow
        sprAutoTeacher.Col = sprAutoTeacher.MaxCols
        
        If sprAutoTeacher.Value = 1 Then        '선택한 것만 처리
        
            sprAutoTeacher.Col = 2:         sTcrCD = Trim(sprAutoTeacher.Text)      '< 강사코드
            sprAutoTeacher.Col = 3:         sSubjCD = Trim(sprAutoTeacher.Text)     '< 과목코드
            
            sprAutoTeacher.Col = 6:         sTcrNM = Trim(sprAutoTeacher.Text)      '< 강사명
            sprAutoTeacher.Col = 7:         sSubjNM = Trim(sprAutoTeacher.Text)     '< 과목명
                sTmrTcrSubj = sSubjNM & "," & sTcrNM                                    '< "과목,강사" : 시간표에서 비교
            
            sprAutoTeacher.Col = 4:         sLsnCD = Trim(sprAutoTeacher.Text)      '< 반코드
            
            sprAutoTeacher.Col = 9:         nSisu = sprAutoTeacher.Value            '< 남은시수
        
            sprAutoTeacher.Col = 5              '< 과목구분
                Select Case Trim(sprAutoTeacher.Text)
                    Case "10", "20", "30"               '# 언 수 외
                        With sprWork
                            For nRowj = 1 To .MaxRows Step 1
                                For nCol = 1 To .MaxCols Step 1
                                    .Row = nRowj
                                    .Col = nCol
                                        .Text = "1"         '<< 1. 은 배정가능 X 는 배정불가
                                Next nCol
                            Next nRowj
                        End With
                        
                        Call Data_MTX01("X", sLsnCD)                '>> 2. 구조별 시간표 내역에서 불가능 부분 추출
                        Call init_Work("1")                         '>> 3. 전체 초기화 (전체를 선택가능 상태로)
                        Call Data_TCR(sTcrCD)                       '>> 4. 배정된 내용 제외 (강사가 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_Lsn(sLsnCD)                       '>> 5. 배정된 내용 제외 (반이 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_not_Teaching(sTcrCD, sSubjCD)     '>> 6. 강의불가능 시수
                        
                        
                    Case "40", "50"                     '# 사 과탐
                        With sprWork
                            For nRowj = 1 To .MaxRows Step 1
                                For nCol = 1 To .MaxCols Step 1
                                    .Row = nRowj
                                    .Col = nCol
                                        .Text = "X"         '<< 1. 은 배정가능 X 는 배정불가
                                Next nCol
                            Next nRowj
                        End With
                        
                        Call Data_MTX01("1", sLsnCD)                '>> 2. 구조별 시간표 내역에서 가능부분 추출
                        Call init_Work("X")                         '>> 3. 전체 초기화 (전체를 선택불가능 상태로)
                        Call Data_TCR(sTcrCD)                       '>> 4. 배정된 내용 제외 (강사가 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_Lsn(sLsnCD)                       '>> 5. 배정된 내용 제외 (반이 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_not_Teaching(sTcrCD, sSubjCD)     '>> 6. 강의불가능 시수
                        
                        
                End Select
                
            
            '----------------------------------------------------------------------------------
            '<< 배정 불가능 시수가 시간표에 모두 배정된 상태임.
            '   "1" 인 부분만 찾아서 데이터 등록하면 됨.
            
                Call Save_Auto_Time_Schedule(sTcrCD, sSubjCD, sLsnCD, nSisu)
                
            '----------------------------------------------------------------------------------
            
            
            sprAutoTeacher.Col = sprAutoTeacher.MaxCols
            sprAutoTeacher.Value = 0                        '< 선택해제
            
        End If
    Next nRow
    
    
    
    '모든 작업이 다 끝나면 다시 조회
    '<< 저장된 데이터를 조회합니다. >>
    Call cmdFind_Click
    Call cmdSearchTcr_Click
    Call cmdCalcu_TCR_Click
    
    
    MsgBox "작업 완료하였습니다.", vbInformation + vbOKOnly, "자동 시간표 배분하기"
    
    
End Sub


'## 탐구영역 자유배정
Private Sub cmdWorkTamgu_Click()
    
    Dim nRow        As Long
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    Dim sLsnCD      As String
    
    Dim sTcrNM      As String
    Dim sSubjNM     As String
    Dim sTmrTcrSubj As String       '< "과목,강사"
    
    Dim nSisu       As Long
    
    Dim nRowj       As Long
    Dim nCol        As Long
    
    If MsgBox("자동배분할 강사를 선택하셨습니까?" & vbCrLf & _
              "선택하였으면 확인을 클릭하십시요." & vbCrLf & _
              "과목별 순서를 정하세요.", vbQuestion + vbYesNo, "강사별 시수현황") = vbNo Then
        Exit Sub
    End If
    
    ProgressBar2.Value = 0
    
    For nRow = 1 To sprAutoTeacher.MaxRows Step 1
    
        ProgressBar2.Value = Fix(nRow / sprAutoTeacher.MaxRows * 100)
        
        sprAutoTeacher.Row = nRow
        sprAutoTeacher.Col = sprAutoTeacher.MaxCols
        
        If sprAutoTeacher.Value = 1 Then        '선택한 것만 처리
        
            sprAutoTeacher.Col = 2:         sTcrCD = Trim(sprAutoTeacher.Text)      '< 강사코드
            sprAutoTeacher.Col = 3:         sSubjCD = Trim(sprAutoTeacher.Text)     '< 과목코드
            
            sprAutoTeacher.Col = 6:         sTcrNM = Trim(sprAutoTeacher.Text)      '< 강사명
            sprAutoTeacher.Col = 7:         sSubjNM = Trim(sprAutoTeacher.Text)     '< 과목명
                sTmrTcrSubj = sSubjNM & "," & sTcrNM                                    '< "과목,강사" : 시간표에서 비교
            
            sprAutoTeacher.Col = 4:         sLsnCD = Trim(sprAutoTeacher.Text)      '< 반코드
            
            sprAutoTeacher.Col = 9:         nSisu = sprAutoTeacher.Value            '< 남은시수
        
            sprAutoTeacher.Col = 5              '< 과목구분
                Select Case Trim(sprAutoTeacher.Text)
                    Case "10", "20", "30"               '# 언 수 외
                        ' NO ACTION
                    Case "40", "50"                     '# 사 과탐
                        With sprWork
                            For nRowj = 1 To .MaxRows Step 1
                                For nCol = 1 To .MaxCols Step 1
                                    .Row = nRowj
                                    .Col = nCol
                                        .Text = "1"         '<< 1. 은 배정가능 X 는 배정불가
                                Next nCol
                            Next nRowj
                        End With
                        
                        Call Data_MTX01("X", sLsnCD)                '>> 2. 구조별 시간표 내역에서 불가능 부분 추출
                        Call init_Work("1")                         '>> 3. 전체 초기화 (전체를 선택가능 상태로)
                        Call Data_TCR(sTcrCD)                       '>> 4. 배정된 내용 제외 (강사가 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_Lsn(sLsnCD)                       '>> 5. 배정된 내용 제외 (반이 요일/교시에 이미 등록된 경우 제외시킴)
                        Call Data_not_Teaching(sTcrCD, sSubjCD)     '>> 6. 강의불가능 시수
                End Select
                
            
            '----------------------------------------------------------------------------------
            '<< 배정 불가능 시수가 시간표에 모두 배정된 상태임.
            '   "1" 인 부분만 찾아서 데이터 등록하면 됨.
            
                Call Save_Auto_Time_Schedule(sTcrCD, sSubjCD, sLsnCD, nSisu)
                
            '----------------------------------------------------------------------------------
            
            
            sprAutoTeacher.Col = sprAutoTeacher.MaxCols
            sprAutoTeacher.Value = 0                        '< 선택해제
            
        End If
    Next nRow
    
    '모든 작업이 다 끝나면 다시 조회
    '<< 저장된 데이터를 조회합니다. >>
    Call cmdFind_Click
    Call cmdSearchTcr_Click
    Call cmdCalcu_TCR_Click
    
    
    MsgBox "작업 완료하였습니다.", vbInformation + vbOKOnly, "자동 시간표 배분하기"
    
End Sub



    
    '## 시간표 배정내역 저장
    Private Sub Save_Auto_Time_Schedule(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aLsnCD As String, ByVal aSisu As Long)
        Dim nRec        As Long
        
        Dim nRow        As Long
        Dim nCol        As Long
        
        Dim sWeek       As String
        Dim sLesson     As String
        
        Dim sStr        As String
        
        Dim DBCmd       As ADODB.Command
        Dim DBParam     As ADODB.Parameter
        Dim ni          As Long
        Dim nExe        As Long
        
        Dim nTotExe     As Long
        
    
        On Error Resume Next
        
        nTotExe = 0
        
        For nRec = 1 To aSisu Step 1
            With sprWork
            
                For nRow = 1 To .MaxRows Step 1
                    For nCol = 1 To .MaxCols Step 1
                        .Row = nRow
                        .Col = nCol
                    
                        If StrComp(Trim(.Text), "1", vbTextCompare) = 0 Then        '< 분배 방법은 요일로 분배됩니다.
                                                                                    '  한번 선택되면 다시 선택하지 않음.
                                                                                    
                            .Text = "A"
                            
                            Select Case .Col                        '<< 요일
                                Case 1
                                    sWeek = "2"       ' 월
                                Case 2
                                    sWeek = "3"
                                Case 3
                                    sWeek = "4"
                                Case 4
                                    sWeek = "5"
                                Case 5
                                    sWeek = "6"
                                Case 6
                                    sWeek = "7"       ' 토
                                Case 7
                                    sWeek = "1"       ' 일
                            End Select
                            sLesson = Trim(CStr(.Row))              '<< 교시
                            
                            
                            basDataBase.DBConn.BeginTrans
                            
                            Set DBCmd = New ADODB.Command
                            Set DBParam = New ADODB.Parameter
                            
                            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                                        


                            
                            nExe = 0
                            
                            '< insert 만 있음 >
                            '<< INSERT >>
                            sStr = ""
                            sStr = sStr & "  INSERT INTO SDTRX50TB ( YM, ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS ) "
                            sStr = sStr & "  VALUES ( "
                            sStr = sStr & "          '" & Trim(fpYM.UnFmtText) & "',"
                            sStr = sStr & "          '" & Trim(basModule.SchCD) & "',"
                            sStr = sStr & "          '" & aTcrCD & "',"
                            sStr = sStr & "          '" & aSubjCD & "',"
                            sStr = sStr & "          '" & aLsnCD & "',"
                            sStr = sStr & "          " & sLesson & ", "
                            sStr = sStr & "          " & sWeek
                            sStr = sStr & "  ) "
                            
                            DBCmd.CommandText = sStr
                            DBCmd.CommandType = adCmdText
                            DBCmd.CommandTimeout = 30
                            
                            DBCmd.Execute nExe, , -1
                                            
                            Do While basDataBase.DBConn.State And adStateExecuting
                                DoEvents
                            Loop
                                    
                            If nExe = 1 Then
                                basDataBase.DBConn.CommitTrans
                                
                                nTotExe = nTotExe + 1
                                If nTotExe = aSisu Then
                                
                                    Set DBCmd = Nothing
                                    Set DBParam = Nothing
                                    
                                    Exit Sub
                                End If
                            Else
                                basDataBase.DBConn.RollbackTrans
                            End If
                            
                        End If
                        
                        
                        
                    Next nCol
                Next nRow
            End With
        Next nRec
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
    End Sub






    '## 6. 강의불가능 시수
    Private Sub Data_not_Teaching(ByVal aTcrCD As String, ByVal aSubjCD As String)
    
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        Dim sStr        As String
        
        Dim sTmp        As String
        Dim nTmp        As Long
        
        Dim ni          As Long
        Dim nRec        As Long
        
        Dim nWeek       As Long
        Dim nLesson     As Long
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT LESSON, WEEKS"
        sStr = sStr & "          From SDTCR15TB"
        sStr = sStr & "         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "           AND TCRCD  = '" & aTcrCD & "'"
        sStr = sStr & "           AND SUBJCD = '" & aSubjCD & "'"
        
        
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
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< 요일
                            Case "2"
                                nWeek = 1
                            Case "3"
                                nWeek = 2
                            Case "4"
                                nWeek = 3
                            Case "5"
                                nWeek = 4
                            Case "6"
                                nWeek = 5
                            Case "7"
                                nWeek = 6
                            Case "1"
                                nWeek = 7
                        End Select
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< 교시
                        
                        sprWork.Row = nLesson
                        sprWork.Col = nWeek
                            sprWork.Text = "X"          '< 불가능 체크 : 이미 등록된 강사이므로 배정불능
                    
                    .MoveNext
                Next nRec
            End If
        End With
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBRec = Nothing
        Set DBParam = Nothing
        
        On Error GoTo 0
    End Sub
  
    
    '## 5. 배정된 내용 제외 (반이 요일/교시에 이미 등록된 경우 제외시킴)
    Private Sub Data_Lsn(ByVal aLsnCD As String)
    
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        Dim sStr        As String
        
        Dim sTmp        As String
        Dim nTmp        As Long
        
        Dim ni          As Long
        Dim nRec        As Long
        
        Dim nWeek       As Long
        Dim nLesson     As Long
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT WEEKS, LESSON"
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N", "J"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "S"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "K"
                sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
        End Select
        
        sStr = sStr & "                       A.TCRCD, A.TCRNM,"
        sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
        sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
        sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                         WHERE A.ACID   = B.ACID"
        sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) A,"
        
        sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
        sStr = sStr & "                           FROM SDLSN01TB "
        sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'        sStr = sStr & "                         UNION"
'        sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
'        sStr = sStr & "                           FROM SDLSN02TB "
'        sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) B"
        
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N", "J"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "S"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "K"
                sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
        End Select
        
        sStr = sStr & "                       A.TCRCD, A.TCRNM ,"
        sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
        sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
        sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                         WHERE A.ACID   = B.ACID"
        sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) A,"
        sStr = sStr & "                       SDLSN02TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT '00000' AS LSNCD, PRT_LSNNM AS LSNNM,"
        sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','인문계','2','자연계','기타') AS KAEYOLNM,"
        sStr = sStr & "                       '' AS CLASSNM,"
        sStr = sStr & "                       '' AS DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        sStr = sStr & "                       PRT_LSN AS LSNCDNM,"
        sStr = sStr & "                       B.TCRCD, B.TCRNM,"
        sStr = sStr & "                       B.SUBJCD, B.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                 WHERE A.ACID   = B.ACID"
        sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                   AND A.LSNCD  = '00000'"
        sStr = sStr & "               )"
        sStr = sStr & "         WHERE LSNCD  = '" & aLsnCD & "'"
        
        
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
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< 요일
                            Case "2"
                                nWeek = 1
                            Case "3"
                                nWeek = 2
                            Case "4"
                                nWeek = 3
                            Case "5"
                                nWeek = 4
                            Case "6"
                                nWeek = 5
                            Case "7"
                                nWeek = 6
                            Case "1"
                                nWeek = 7
                        End Select
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< 교시
                        
                        sprWork.Row = nLesson
                        sprWork.Col = nWeek
                            sprWork.Text = "X"          '< 불가능 체크 : 이미 등록된 강사이므로 배정불능
                    
                    .MoveNext
                Next nRec
            End If
        End With
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBRec = Nothing
        Set DBParam = Nothing
        
        On Error GoTo 0
    End Sub



    '## 4. 배정된 내용 제외 (과목/강사가 요일/교시에 이미 등록된 경우 제외시킴)
    Private Sub Data_TCR(ByVal aTcrCD As String)
    
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        Dim sStr        As String
        
        Dim sTmp        As String
        Dim nTmp        As Long
        
        Dim ni          As Long
        Dim nRec        As Long
        
        Dim nWeek       As Long
        Dim nLesson     As Long
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT WEEKS, LESSON"
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N", "J"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "S"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "K"
                sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
        End Select

        sStr = sStr & "                       A.TCRCD, A.TCRNM,"
        sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
        sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
        sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                         WHERE A.ACID   = B.ACID"
        sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) A,"
        
        sStr = sStr & "                        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
        sStr = sStr & "                           FROM SDLSN01TB "
        sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'        sStr = sStr & "                         UNION"
'        sStr = sStr & "                         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
'        sStr = sStr & "                           FROM SDLSN02TB "
'        sStr = sStr & "                          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) B"
        
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N", "J"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "S"
                sStr = sStr & "               SUBSTR(B.KAEYOL,2,1)||B.LSNCDNM AS LSNCDNM,"
            Case "K"
                sStr = sStr & "               SUBSTR(A.SUBJNM,1,1)||B.LSNCDNM AS LSNCDNM,"
        End Select
        
        sStr = sStr & "                       A.TCRCD, A.TCRNM ,"
        sStr = sStr & "                       A.SUBJCD, A.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
        sStr = sStr & "                               B.TCRNM, B.SUBJNM, B.TCRCD, B.SUBJCD"
        sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                         WHERE A.ACID   = B.ACID"
        sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                           AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                        ) A,"
        sStr = sStr & "                       SDLSN02TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT '00000' AS LSNCD, PRT_LSNNM AS LSNNM,"
        sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','인문계','2','자연계','기타') AS KAEYOLNM,"
        sStr = sStr & "                       '' AS CLASSNM,"
        sStr = sStr & "                       '' AS DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        sStr = sStr & "                       PRT_LSN AS LSNCDNM,"
        sStr = sStr & "                       B.TCRCD, B.TCRNM,"
        sStr = sStr & "                       B.SUBJCD, B.SUBJNM,"
        sStr = sStr & "                       A.WEEKS, A.LESSON"
        sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
        sStr = sStr & "                 WHERE A.ACID   = B.ACID"
        sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
        sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
        sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
        sStr = sStr & "                   AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                   AND A.LSNCD  = '00000'"
        sStr = sStr & "               )"
        sStr = sStr & "         WHERE TCRCD  = '" & aTcrCD & "'"
        
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
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< 요일
                            Case "2"
                                nWeek = 1
                            Case "3"
                                nWeek = 2
                            Case "4"
                                nWeek = 3
                            Case "5"
                                nWeek = 4
                            Case "6"
                                nWeek = 5
                            Case "7"
                                nWeek = 6
                            Case "1"
                                nWeek = 7
                        End Select
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< 교시
                        
                        sprWork.Row = nLesson
                        sprWork.Col = nWeek
                            sprWork.Text = "X"          '< 불가능 체크 : 이미 등록된 강사이므로 배정불능
                    
                    .MoveNext
                Next nRec
            End If
        End With
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBRec = Nothing
        Set DBParam = Nothing
        
        On Error GoTo 0
    End Sub

    '## 3. 전체 초기화
    Private Sub init_Work(ByVal ainitVal As String)
        Dim nRow        As Long
        Dim nCol        As Long
        
        Dim nLesson     As Long
        Dim nWeeks      As Long
        
        With sprWork
            
            Select Case Trim(txtWeeks.Text)
                Case "월", "화", "수", "목", "금"
                    nWeeks = 6
                Case "토"
                    nWeeks = 7
                Case "일"
                    nWeeks = 8
            End Select
            
            Select Case fpLesson.Value
                Case Is <= 7
                    nLesson = 8
                Case Is = 8
                    nLesson = 9
                Case Is = 9
                    nLesson = 10
                Case Is = 10
                    nLesson = 11
            End Select
            
            For nRow = nLesson To .MaxRows Step 1
                For nCol = 1 To .MaxCols Step 1
                    .Row = nRow
                    .Col = nCol
                        .Text = "X"                 '< 시간표 배정없음
                Next nCol
            Next nRow
            
            For nCol = nWeeks To .MaxCols Step 1
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = nCol
                        .Text = "X"                 '< 시간표 배정없음
                Next nRow
            Next nCol
            
        End With
    End Sub

    '## 2. 구조별 시간표 내역 조회
    Private Sub Data_MTX01(ByVal aAlloc As String, ByVal aLsnCD As String)
        
        Dim sLsnType    As String
        
        Dim DBCmd       As ADODB.Command
        
        Dim DBParam     As ADODB.Parameter
        Dim DBRec       As ADODB.Recordset
        Dim DBRecj      As ADODB.Recordset
        
        Dim nLength     As Long
        Dim sStr        As String
    
        Dim sTmp        As String
        Dim nTmp        As Long
    
        Dim ni          As Long
        Dim nRec        As Long
        Dim nRecj       As Long
        
        Dim nLesson     As Long
        Dim nWeek       As Long
        
        Dim nRow        As Long
        Dim nCol        As Long
        
        On Error GoTo ErrStmt
    
        sStr = ""
        sStr = sStr & "        SELECT A.ACID, A.KAEYOL, A.LSNTYPE, A.LSNCD"
        sStr = sStr & "          FROM SDLSN06TB A, "
        
        sStr = sStr & "               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                      '2009.01.12 추가
        sStr = sStr & "                  FROM SDLSN01TB "
        sStr = sStr & "                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION"
        sStr = sStr & "                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
        sStr = sStr & "                  FROM SDLSN02TB "
        sStr = sStr & "                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "               ) B"
        
        sStr = sStr & "         Where A.ACID  = B.ACID"
        sStr = sStr & "           AND A.LSNCD = B.LSNCD"
        sStr = sStr & "           AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "           AND A.LSNCD BETWEEN '00001' AND '89999'"
        sStr = sStr & "           AND A.LSNCD = '" & aLsnCD & "'"
        sStr = sStr & "         GROUP BY A.ACID, A.KAEYOL, A.LSNTYPE, A.LSNCD"

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
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        If DBRec.RecordCount < 1 Then
            '## 구조별 시간표 등록 내용 없음.
            
        Else
            DBRec.MoveFirst
               
            For nRec = 1 To DBRec.RecordCount Step 1
            
                Set DBRecj = New ADODB.Recordset
            
                sStr = ""
                Select Case aAlloc
                    Case "X"            '< 시간배정 불가능 부분 추출
                        sStr = sStr & "        SELECT KAEYOL, LESSON, WEEKS"
                        sStr = sStr & "          FROM (SELECT KAEYOL, LESSON, WEEKS"
                        sStr = sStr & "                  From SDTRX11TB"
                        sStr = sStr & "                 WHERE ACID   =    '" & Trim(basModule.SchCD) & "'"
                        sStr = sStr & "                   AND TRXCD  LIKE '" & Trim(DBRec.Fields("LSNTYPE")) & "%'"
                        sStr = sStr & "                   AND KAEYOL =    '" & Trim(DBRec.Fields("KAEYOL")) & "'"
                        sStr = sStr & "                Union All"
                        sStr = sStr & "                SELECT KAEYOL, LESSON, WEEKS"
                        sStr = sStr & "                  From SDTRX11TB"
                        sStr = sStr & "                 WHERE ACID   =    '" & Trim(basModule.SchCD) & "'"
                        sStr = sStr & "                   AND TRXCD  LIKE 'PB%' "
                        sStr = sStr & "                   AND KAEYOL =    '" & Trim(DBRec.Fields("KAEYOL")) & "'"
                        sStr = sStr & "                )"
                        
                    Case "1"
                        sStr = sStr & "        SELECT KAEYOL, LESSON, WEEKS"
                        sStr = sStr & "          From SDTRX11TB"
                        sStr = sStr & "         WHERE ACID   =    '" & Trim(basModule.SchCD) & "'"
                        sStr = sStr & "           AND TRXCD  LIKE '" & Trim(DBRec.Fields("LSNTYPE")) & "%'"
                        sStr = sStr & "           AND KAEYOL =    '" & Trim(DBRec.Fields("KAEYOL")) & "'"
                                                
                End Select
                
                
                DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
                
                
 
 
                
                DBRecj.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
                Do While DBRecj.State And adStateExecuting
                    DoEvents
                Loop
                
                If DBRecj.RecordCount < 1 Then
                    'NOTHING
                Else
                    DBRecj.MoveFirst
                    For nRecj = 1 To DBRecj.RecordCount Step 1
                        Select Case Trim(DBRecj.Fields("WEEKS"))    '< 요일
                            Case "2"
                                nWeek = 1
                            Case "3"
                                nWeek = 2
                            Case "4"
                                nWeek = 3
                            Case "5"
                                nWeek = 4
                            Case "6"
                                nWeek = 5
                            Case "7"
                                nWeek = 6
                            Case "1"
                                nWeek = 7
                        End Select
                        nLesson = CLng(DBRecj.Fields("LESSON"))     '< 교시
                        
                        sprWork.Row = nLesson
                        sprWork.Col = nWeek
                            sprWork.Text = aAlloc       '< 가능여부 파악
                        
                        DBRecj.MoveNext
                    Next nRecj
                End If
                
                DBRec.MoveNext
            Next nRec
        End If
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBParam = Nothing
        Set DBRec = Nothing
        Set DBRecj = Nothing
    
        On Error GoTo 0
    End Sub






































'--------------------------------------------------------------
' 송장 엑셀자료 만들기
'--------------------------------------------------------------
Private Sub cmdExcel_Click()
    Call Make_Tmr_ExcelFile
End Sub

Private Sub Make_Tmr_ExcelFile()
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim sComp       As String
    
    Dim sFileName   As String
    Dim sFilePath   As String
    Dim sLogFile    As String
    
    Dim nWeekSrt    As Long
    Dim nColor      As Long
    
    Dim nRet        As Long
    Dim nRow2       As Long
    
    
    Dim sTcrTmp     As String
    Dim sTcrComp    As String
    Dim nChkRow     As Long
    
    Dim sTSisu      As String
    Dim sSSisu      As String
    
    If sprTmr_Tcr.MaxRows = 0 Then Exit Sub
    
    If MsgBox("시간표 엑셀자료 생성하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrDlg

    If Dir(App.Path & "\TMR_EXCEL", vbDirectory) = "" Then MkDir App.Path & "\TMR_EXCEL"

    'TEXT파일을 생성 처리합니다.
    With dlgExcel
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path & "\TMR_EXCEL"
        .Filter = "DAT FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowSave

        '파일명을 처리합니다.
        If (.fileName) = "" Then Exit Sub
        
        sFileName = Mid$(dlgExcel.FileTitle, 1, InStr(1, dlgExcel.FileTitle, ".", vbTextCompare) - 1)
        sFilePath = Mid$(dlgExcel.fileName, 1, Len(dlgExcel.fileName) - InStrB(1, dlgExcel.fileName, "\", vbTextCompare) - 1)
        sLogFile = sFilePath & sFileName & ".LOG"
        
    End With

    On Error GoTo 0
    On Error GoTo ErrExcel
    
    sprExcel.ColHeadersShow = True
    sprExcel.RowHeadersShow = True
    
    sprExcel.MaxRows = 0
    sprExcel.MaxCols = 0
    
    For nRow = 1 To sprTmr_Tcr.ColHeaderRows Step 1
        sprTmr_Tcr.Row = SpreadHeader + nRow - 1
            '< 데이타 복사 >
            sprExcel.MaxRows = sprExcel.MaxRows + 1
            sprExcel.Row = sprExcel.MaxRows                                         '< header row
        
            sprExcel.MaxCols = sprTmr_Tcr.RowHeaderCols + sprTmr_Tcr.MaxCols        '< 전체 cols
        
            '< Row Header 생성 >
            For nCol = 1 To sprTmr_Tcr.RowHeaderCols Step 1
                sprTmr_Tcr.Col = SpreadHeader + nCol - 1
                    sTmp = sprTmr_Tcr.Text
                    
                    sprExcel.Col = nCol                                                 '< 데이터 넣음
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
                    
                    With sprExcel
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.ShadowColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End With
            Next nCol
            
            '< Data >
            For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
                sprTmr_Tcr.Col = nCol
                    sTmp = Trim(sprTmr_Tcr.Text)
                
                    sprExcel.Col = sprTmr_Tcr.RowHeaderCols + nCol
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
            Next nCol
            
            sprExcel.SetCellBorder sprTmr_Tcr.RowHeaderCols + 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor1, CellBorderStyleSolid
            
            With sprExcel
                .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.ShadowColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
    Next nRow
    
    '< Data 부분 >
    For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
        sprTmr_Tcr.Row = nRow
        sprTmr_Tcr.Col = SpreadHeader:      sTcrComp = Trim(sprTmr_Tcr.Text)
        
        nChkRow = 0
        For nRow2 = 1 To sprExcel.MaxRows Step 1
            sprExcel.Row = nRow2
            sprExcel.Col = 1:       sTcrTmp = Trim(sprExcel.Text)
            
            If StrComp(sTcrComp, sTcrTmp, vbTextCompare) = 0 Then
                nChkRow = nRow2
                Exit For
            End If
        Next nRow2
        
        If nChkRow = 0 Then         '<- 1줄 추가
            '< 데이타 복사 >
            sprExcel.MaxRows = sprExcel.MaxRows + 1
            sprExcel.Row = sprExcel.MaxRows                                         '< header row
            
            '< Row Header 생성 >
            For nCol = 1 To sprTmr_Tcr.RowHeaderCols Step 1
                sprTmr_Tcr.Col = SpreadHeader + nCol - 1
                    sTmp = sprTmr_Tcr.Text
                    
                    sprExcel.Col = nCol                                                 '< 데이터 넣음
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    sprExcel.ColWidth(sprExcel.Col) = 5
            Next nCol
            
        Else                        '<- 해당열에 추가
            sprExcel.Row = nChkRow
            
            '< Row Header 생성 >
            For nCol = 1 To sprTmr_Tcr.RowHeaderCols Step 1
                sprTmr_Tcr.Col = SpreadHeader + nCol - 1
                    
                    Select Case nCol
                        Case 6, 7
                            sTmp = sprTmr_Tcr.Text
                            
                            If IsNumeric(sTmp) = True Then
                                sprExcel.Col = nCol
                                sTSisu = sprExcel.Text
                                If IsNumeric(sTSisu) = True Then
                                    sTmp = Trim(CStr(CLng(sTSisu) + CLng(sTmp)))
                                End If
                            End If
                            sprExcel.Col = nCol                                                 '< 데이터 넣음
                                Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                                
                    End Select
            Next nCol
            
        End If
        
        '< Data >
        For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
            sprTmr_Tcr.Col = nCol
                If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                   sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                    nColor = sprTmr_Tcr.BackColor
                Else
                    nColor = basModule.WhiteColor
                End If
                
                sTmp = Trim(sprTmr_Tcr.Text)
                
                sprExcel.Col = sprTmr_Tcr.RowHeaderCols + nCol
                sprExcel.ColWidth(sprExcel.Col) = 3
                
                If Trim(sTmp) = "" Then
                    ' no action
                    If nColor = lblNotTeaching.BackColor Then
                        If Trim(sprExcel.Text) <> "#" Then
                            sTmp = "#" & Trim(sprExcel.Text)
                            Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        End If
                    End If
                Else
                    If Trim(sprExcel.Text) <> "" Then
                        If StrComp(Trim(sprExcel.Text), "#", vbTextCompare) = 0 Then
                            sTmp = "#" & sTmp
                        Else
                            If nColor = lblNotTeaching.BackColor Then
                                sTmp = "#" & sTmp
                            Else
                                sTmp = sTmp & "/" & Trim(sprExcel.Text)
                            End If
                        End If
                        Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    Else
                        If nColor = lblNotTeaching.BackColor Then
                            sTmp = "#" & sTmp
                        End If
                        Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    End If
                End If
                
                sprExcel.Row2 = sprExcel.Row
                sprExcel.Col2 = sprExcel.Col
                sprExcel.BlockMode = True
                    sprExcel.BackColor = nColor
                    sprExcel.BackColorStyle = BackColorStyleUnderGrid
                sprExcel.BlockMode = False
        Next nCol
            
        If sprExcel.Row > 4 Then
            If ((sprExcel.Row - sprTmr_Tcr.ColHeaderRows) Mod 5) = 0 Then sprExcel.SetCellBorder 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
        End If
        
        With sprExcel
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = sprTmr_Tcr.RowHeaderCols
            .BlockMode = True
                .BackColor = basModule.ShadowColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End With
    
    Next nRow
        
    '< 데이터 정렬 및 맞춤 >
    With sprExcel
        If .MaxRows > 1 Then
            .SetCellBorder 1, 1, .MaxCols, .MaxRows, 16, &H80000008, CellBorderStyleSolid

            .Row = 3
                .SetCellBorder 1, .Row, .MaxCols, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid

            .Row = 2
            nWeekSrt = 0
            For nCol = 1 To .MaxCols Step 1
                .Col = nCol:    sTmp = Trim(.Text):     If sComp = "" Then sComp = Trim(.Text)
                If Trim(.Text) <> "" Then
                    If StrComp(sComp, sTmp, vbTextCompare) <> 0 Then
                        .SetCellBorder .Col, 1, .Col, .MaxRows, 1, basModule.SectionColor2, CellBorderStyleSolid
                        sComp = sTmp

                        If nWeekSrt = 0 Then
                            .AddCellSpan sprTmr_Tcr.RowHeaderCols + 1, 1, nCol - sprTmr_Tcr.RowHeaderCols - 1, 1
                        Else
                            .AddCellSpan nWeekSrt, 1, nCol - nWeekSrt, 1
                        End If
                        nWeekSrt = nCol

                    End If
                End If
            Next nCol
            If nWeekSrt = 0 Then
                .AddCellSpan sprTmr_Tcr.RowHeaderCols + 1, 1, nCol - sprTmr_Tcr.RowHeaderCols - 1, 1
            Else
                .AddCellSpan nWeekSrt, 1, .MaxCols - nWeekSrt + 1, 1
            End If
            nWeekSrt = nCol

            .Row = 2:       .DeleteRows .Row, 1:        .MaxRows = .MaxRows - 1

            .Col = 7
                .SetCellBorder .Col, 1, .Col, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        End If
    End With
    
    nRet = sprExcel.ExportToExcel(dlgExcel.fileName, "Time_Schedule", sLogFile)
    
    MsgBox "엑셀작성하였습니다." & vbCrLf & _
           "확인하십시요.", vbInformation + vbOKOnly, "시간표 엑셀자료 만들기"

    Exit Sub
ErrExcel:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 엑셀자료 만들기"
    Exit Sub
ErrDlg:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성을 취소하였습니다.", vbCritical + vbOKOnly, "엑셀자료 생성"
End Sub











'## 전체 시간표 삭제
Private Sub cmdTmrAllDelete_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim nExe        As Long
    Dim ni          As Long
    Dim sStr        As String
    
    If MsgBox("작성한 전체 시간표를 삭제하시겠습니까?" & vbCrLf & _
              "전체를 삭제하시면 처음부터 작업을 다시 하셔야하니 유의하십시요.", vbQuestion + vbYesNo, "시간표 삭제") = vbNo Then
         Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    

    '< DELETE 만 있음 >

    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDTRX50TB "
    sStr = sStr & "  WHERE YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND ACID = '" & Trim(basModule.SchCD) & "'"
                            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
                                    
    If nExe >= 1 Then
        basDataBase.DBConn.CommitTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Call cmdFind_Click
        Call cmdSearchTcr_Click
        
        MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간표 삭제"
    Else
        basDataBase.DBConn.RollbackTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        MsgBox "삭제할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 삭제"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "삭제중 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 삭제"
    basDataBase.DBConn.RollbackTrans
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
            
    On Error GoTo 0

End Sub

'## 언수외 내역만 삭제
Private Sub cmdDelKME_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim nExe        As Long
    Dim ni          As Long
    Dim sStr        As String
    
    If MsgBox("작성한 시간표 중 언/수/외 과목강사의 내용만 삭제하시겠습니까?" & vbCrLf & _
              "삭제는【 강사 및 강사별 과목넣기 】에 과목구분으로 삭제대상을 선정합니다.", vbQuestion + vbYesNo, "시간표 언/수/외 과목내역 삭제") = vbNo Then
         Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection



    '< DELETE 만 있음 >
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDTRX50TB "
    sStr = sStr & "  WHERE YM = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND (ACID, TCRCD, SUBJCD)"
    sStr = sStr & "     IN (SELECT ACID, TCRCD, SUBJCD"
    sStr = sStr & "           From SDTCR01TB"
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            AND SUBJGBN IN ('10','20','30') "                '< 과목구분이 언/ 수/ 외
    sStr = sStr & "         )"
                            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
                                    
    If nExe >= 1 Then
        basDataBase.DBConn.CommitTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Call cmdFind_Click
        Call cmdSearchTcr_Click
        
        MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "시간표 삭제"
    Else
        basDataBase.DBConn.RollbackTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        MsgBox "삭제할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 삭제"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "삭제중 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 삭제"
    basDataBase.DBConn.RollbackTrans
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
            
    On Error GoTo 0
End Sub






'------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------

'############################################### 시간표 배분작업 ################################################################################
Private Sub cmdTmrChg_Click()
    Load TMR052
    TMR052.Show
    TMR052.ZOrder 0
    
End Sub






'############################################### 강의불가시간 보기 ################################################################################
Private Sub cmdViewNotTeach_Click()

    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim sTmp        As String
    
    Dim ni          As Long
    
    Dim nRec        As Long
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim stTcrCD     As String
    Dim stSubjCD    As String
    Dim stWeek      As String
    Dim stLesson    As String
    
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    '> 색 초기화
    With sprTmr_Tcr
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow:    .Col = nCol
                
                If Trim(.Text) = "" Then
                    .Row2 = .Row:   .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                Else
                    If .BackColor = basModule.SectionColor1 Then
                        'NO ACTION
                    Else
                        If .BackColor = lblNotTeaching.BackColor Then
                            'NO ACTION
                        Else
                            .Row2 = .Row:   .Col2 = .Col
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                        End If
                    End If
                End If
            Next nCol
        Next nRow
    End With
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT TCRCD, SUBJCD, WEEKS, LESSON"
    sStr = sStr & "   FROM (SELECT A.TCRCD, A.SUBJCD, NVL(WEEKS,0) AS WEEKS, NVL(LESSON,0) AS LESSON"
    sStr = sStr & "           FROM SDTCR01TB A, SDTCR15TB B"
    sStr = sStr & "          WHERE A.ACID   = B.ACID(+)"
    sStr = sStr & "            AND A.TCRCD  = B.TCRCD (+)"
    sStr = sStr & "            AND A.SUBJCD = B.SUBJCD (+)"
    sStr = sStr & "            AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         )"
    sStr = sStr & "  Where LESSON > 0"
    sStr = sStr & "    AND WEEKS  > 0"
    sStr = sStr & "  ORDER BY TCRCD"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            ProgressBar1.Min = 0
            ProgressBar1.Max = 100
            ProgressBar1.Value = 0
        
            For nRec = 1 To .RecordCount Step 1
                
                ProgressBar1.Value = Fix(nRec / .RecordCount * 100)
                
                sTcrCD = "":    If IsNull(.Fields("TCRCD")) = False Then sTcrCD = Trim(.Fields("TCRCD"))
                sSubjCD = "":   If IsNull(.Fields("SUBJCD")) = False Then sSubjCD = Trim(.Fields("SUBJCD"))
                sWeek = "":     If IsNull(.Fields("WEEKS")) = False Then sWeek = Trim(.Fields("WEEKS"))
                sLesson = "":   If IsNull(.Fields("LESSON")) = False Then sLesson = Trim(.Fields("LESSON"))
                    
                nr_Chk = 0
                nc_Chk = 0
                
                For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
                    sprTmr_Tcr.Row = nRow:      nr_Chk = nRow
                        sprTmr_Tcr.Col = SpreadHeader:      stTcrCD = Trim(sprTmr_Tcr.Text)
                        sprTmr_Tcr.Col = SpreadHeader + 1:  stSubjCD = Trim(sprTmr_Tcr.Text)
                                                
                    '> 강사만 맞으면 표시함. : 노량진 요청사항
                    'If StrComp(sTcrCD, stTcrCD, vbTextCompare) = 0 And _
                    '   StrComp(sSubjCD, stSubjCD, vbTextCompare) = 0 Then
                    If StrComp(sTcrCD, stTcrCD, vbTextCompare) = 0 Then
                    
                        For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
                        
                            sprTmr_Tcr.Col = nCol:      nc_Chk = nCol
                            sprTmr_Tcr.Row = SpreadHeader + 1:      stWeek = Trim(sprTmr_Tcr.Text)
                            sprTmr_Tcr.Row = SpreadHeader + 2:      stLesson = Trim(sprTmr_Tcr.Text)
                            
                            '> 요일과 교시가 맞는경우
                            If StrComp(sWeek, stWeek, vbTextCompare) = 0 And _
                               StrComp(sLesson, stLesson, vbTextCompare) = 0 Then
                                                           
                                '모든 조건이 맞음
                                sprTmr_Tcr.Row = nr_Chk:        sprTmr_Tcr.Row2 = sprTmr_Tcr.Row
                                sprTmr_Tcr.Col = nc_Chk:        sprTmr_Tcr.Col2 = sprTmr_Tcr.Col
                                sprTmr_Tcr.BlockMode = True
                                    sprTmr_Tcr.BackColor = lblNotTeaching.BackColor
                                    sprTmr_Tcr.BackColorStyle = BackColorStyleUnderGrid
                                sprTmr_Tcr.BlockMode = False
                            
                            End If
                        Next nCol
                    End If
                Next nRow
                
                .MoveNext
                
            Next nRec
        End If
    End With
    
    'MsgBox "강사별 제약조건을 조회하였습니다.", vbInformation + vbOKOnly, "강사별 제약조건 조회"
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    On Error GoTo 0

End Sub




















'##==============================================================================================##
'## 등록
'##==============================================================================================##
Private Sub cmdIpruck_Click()
    
    Dim nRow        As Long
    Dim nCol        As Long
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    Dim sLsnCD      As String
    Dim sLsnCDNM    As String
    Dim sWeek       As String
    Dim sLesson     As String
    
    Dim sTcrNM      As String
    Dim sSubjNM     As String
    
    Dim sTmpLsn     As String
    Dim sTmpKaeyol  As String
    
    Dim sDivComma() As String
    Dim sSlash()    As String
    
    Dim sPrt_Kaeyol As String
    Dim sPrt_Lsn    As String
    Dim sPrt_LsnNM  As String
    
    Dim sTmp        As String
    
    If Trim(txtinRow.Text) = "" Or _
       Trim(txtinCol.Text) = "" Or _
       Trim(txtData.Text) = "" Then
        MsgBox "등록할 내용이 없습니다.", vbExclamation + vbOKOnly, "등록"
        Exit Sub
    End If
    
    '>> 변경
    '1. 등록된 내용 찾음. <위치>
        'spread 1
    txtData.Text = UCase(txtData.Text)
        
        If StrComp(Trim(txtinSpr.Text), "sprTmr_Lsn", vbTextCompare) = 0 Then
            With sprTmr_Lsn
                .Row = CLng(txtinRow.Text)
                    .Col = SpreadHeader + 1:        sWeek = Trim(.Text)
                    .Col = SpreadHeader + 2:        sLesson = Trim(.Text)
                    
                .Col = CLng(txtinCol.Text)
                    .Row = SpreadHeader + 1:        sLsnCD = Trim(.Text)
                    .Row = SpreadHeader + 3:        sLsnCDNM = Right(Trim(.Text), 1)
                    .Row = SpreadHeader + 2:        sLsnCDNM = sLsnCDNM & Trim(.Text)
                    
                '2. 등록된 내용 삭제
                    '변경내용 display
                
                If Find_TCR_and_Del_TCR(sLsnCD, sWeek, sLesson) = True Then
                    .Row = CLng(txtinRow.Text)
                    .Col = CLng(txtinCol.Text)
                        .Text = ""
                        
                    If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                       sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                        
                    Else
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
                '3. 등록된 내용 저장
                    '변경내용 display
                    If InStr(1, Trim(txtData.Text), "/", vbTextCompare) = 0 Then
                        
                        If InStr(1, Trim(txtData.Text), ",", vbTextCompare) > 0 Then       ' 반드시 【 , 】 로 과목 및 강사로 구분
                            sDivComma() = Split(UCase(Trim(txtData.Text)), ",", -1, vbTextCompare)
                            
                            sSubjNM = Trim(sDivComma(0))        '< 과목명
                            sTcrNM = Trim(sDivComma(1))         '< 강사명
                            
                            'sWeek
                            'sLesson
                            'sLsnCD
                            
                            sTcrCD = "":        sSubjCD = ""
                            Call Find_Tcr_and_Subj_Code(sTcrCD, sSubjCD, sTcrNM, sSubjNM)
                            
                        '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
                            If sTcrCD <> "" And sSubjCD <> "" Then
                            
                            '** 시간표 내역 등록하기 **
                                Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                Call Show_TMR_Tcr(sLsnCD, sWeek, sLesson)
                                Call sprTmr_Lsn_Click(CLng(txtinCol.Text), CLng(txtinRow.Text))
                                
                            End If
                        End If
                    
                    
                    Else        '<< 기타반 [ 과목, 강사/ 계열 / 반명 ]
                        
                        sSlash = Split(UCase(Trim(txtData.Text)), "/", -1, vbTextCompare)
                        
                        sTmp = Trim(sSlash(0))          '<< 과목, 강사
                        
                        If UBound(sSlash) >= 2 Then
                            
                            sDivComma() = Split(UCase(sSlash(0)), ",", -1, vbTextCompare)
                            
                            sSubjNM = Trim(sDivComma(0))        '< 과목명
                            sTcrNM = Trim(sDivComma(1))         '< 강사명
                            
                            'sWeek
                            'sLesson
                            'sLsnCD
                            
                            sTcrCD = "":        sSubjCD = ""
                            Call Find_Tcr_and_Subj_Code(sTcrCD, sSubjCD, sTcrNM, sSubjNM)
                            
                            '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
                            If sTcrCD <> "" And sSubjCD <> "" Then
                                
                                sPrt_Kaeyol = sSlash(1)
                                sPrt_Lsn = Right(sPrt_Kaeyol, 1) & "ZZ"
                                sPrt_LsnNM = sSlash(2)
                                
                                Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD, sPrt_Kaeyol, sPrt_Lsn, sPrt_LsnNM)
                                Call Show_TMR_Tcr(sLsnCD, sWeek, sLesson)
                                Call sprTmr_Lsn_Click(CLng(txtinCol.Text), CLng(txtinRow.Text))
                                
                            End If
                        End If
                        
                    End If
                    
                    
                End If
            End With
        End If
    
        'spread 2
        If StrComp(Trim(txtinSpr.Text), "sprTmr_Tcr", vbTextCompare) = 0 Then
            With sprTmr_Tcr
                .Row = CLng(txtinRow.Text)
                    .Col = SpreadHeader:            sTcrCD = Trim(.Text)
                    .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
                    
                .Col = CLng(txtinCol.Text)
                    .Row = SpreadHeader + 1:        sWeek = Trim(.Text)
                    .Row = SpreadHeader + 2:        sLesson = Right(Trim(.Text), 1)
                    
                sTmpLsn = Right(Left(Trim(txtData.Text), 3), 2)
                sTmpKaeyol = "0" & Left(Trim(txtData.Text), 1)
                                
                Select Case sTmpKaeyol
                    Case "01"
                        Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)            '< LSNCD 나옴
                        
                        sPrt_Kaeyol = ""
                        sPrt_Lsn = ""
                        sPrt_LsnNM = ""
                    Case "02"
                        Call Get_LsnCD_Data(sLsnCD, sTmpKaeyol, sTmpLsn)            '< LSNCD 나옴
                        
                        sPrt_Kaeyol = ""
                        sPrt_Lsn = ""
                        sPrt_LsnNM = ""
                    Case Else
                        MsgBox "등록할 수 없습니다.", vbExclamation + vbOKOnly, "등록"
                        Exit Sub
                End Select
                          
                
                
                
                '2. 등록된 내용 삭제
                    '변경내용 display
                    
                If Find_TCR_and_Del_LSN(sTcrCD, sSubjCD, sLsnCD, sWeek, sLesson) = True Then
                    .Row = CLng(txtinRow.Text)
                    .Col = CLng(txtinCol.Text)
                        .Text = ""
                    
                    If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                       sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                        'no action
                    Else
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
                    '3. 등록된 내용 저장
                    '변경내용 display
                    If InStr(1, Trim(txtData.Text), "/", vbTextCompare) = 0 Then
                            
                        'sWeek
                        'sLesson
                        'sLsnCD
                        'sTcrCD
                        'sSubjCD
                            
                        '> 조회된 강사 및 과목내역이 있어야 함. ------------------------------------------------
                            If sTcrCD <> "" And sSubjCD <> "" Then
                            
                            '** 시간표 내역 등록하기 **
                                Call Save_TMR_Data(sTcrCD, sSubjCD, sWeek, sLesson, sLsnCD)
                                Call Show_TMR_Tcr_inData(sTcrCD, sSubjCD, sLsnCD, sWeek, sLesson)
                                
                                Call sprTmr_Tcr_Click(CLng(txtinCol.Text), CLng(txtinRow.Text))
                                
                            End If
                        
                    End If
                End If
                
            End With
        End If
        
End Sub


'## 전체 시간표 내역에서 보여주기
Public Sub Show_TMR_Tcr_inData(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String

    Dim nRec        As Long
    Dim ni          As Long
    Dim sData       As String

    Dim nRow        As Long
    Dim nCol        As Long

    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sTmpTcrCD   As String
    Dim sTmpSubjCD  As String
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long

    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD, GET_KEAYOL_N_LSN_TCR01(A.ACID, A.LSNCD) AS DS"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.TCRCD  = '" & aTcrCD & "'"
    sStr = sStr & "    AND A.SUBJCD = '" & aSubjCD & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & aLsnCD & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    
    DBRec.MoveFirst
    For nRec = 1 To DBRec.RecordCount Step 1
        
        If IsNull(DBRec.Fields("TCRCD")) = False And _
           IsNull(DBRec.Fields("SUBJCD")) = False And _
           IsNull(DBRec.Fields("DS")) = False Then
            
            sTcrCD = Trim(DBRec.Fields("TCRCD"))
            sSubjCD = Trim(DBRec.Fields("SUBJCD"))
            sData = Trim(DBRec.Fields("DS"))
            
            
            With sprTmr_Tcr
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow:        nChkRow = .Row
                    .Col = SpreadHeader:            sTmpTcrCD = Trim(.Text)
                    .Col = SpreadHeader + 1:        sTmpSubjCD = Trim(.Text)
                    
                    If StrComp(sTcrCD, sTmpTcrCD, vbTextCompare) = 0 And _
                       StrComp(sSubjCD, sTmpSubjCD, vbTextCompare) = 0 Then
                       
                        For nCol = 1 To .MaxCols Step 1
                            .Col = nCol:        nChkCol = .Col
                            .Row = SpreadHeader + 1:        sTmpWeek = Trim(.Text)
                            .Row = SpreadHeader + 2:        sTmpLesson = Trim(.Text)
                            
                            If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                               StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                               
                                .Row = nChkRow
                                .Col = nChkCol
                                
                                If Trim(.Text) = "" Then
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(sprTmr_Tcr, "center", "left", 60, sData)
                                        
                                        If InStr(1, sData, "/", vbTextCompare) > 0 Then
                                            .Row2 = .Row
                                            .Col2 = .Col
                                            .BlockMode = True
                                                .BackColor = basModule.SectionColor1
                                                .BackColorStyle = BackColorStyleUnderGrid
                                            .BlockMode = False
                                            
                                        End If
                                    End If
                                End If
                            End If
                        Next nCol
                    End If
                Next nRow
            End With
        End If
        
        DBRec.MoveNext
    Next nRec
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
        
End Sub




'## 반, 과목, 강사로 삭제
Private Function Find_TCR_and_Del_LSN(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String) As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim bRet        As Boolean
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTmpTcr     As String
    Dim sTmpSubj    As String
    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim nExe        As Long
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long
    
    
    On Error GoTo ErrStmt
    
    bRet = True
    
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.TCRCD  = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "    AND A.SUBJCD = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    If DBRec.RecordCount > 0 Then
        DBRec.MoveFirst
            
        If DBRec.RecordCount > 0 Then
            basDataBase.DBConn.BeginTrans
        End If
            
        For nRec = 1 To DBRec.RecordCount Step 1
            
            sStr = ""
            sStr = sStr & " DELETE"
            sStr = sStr & "   FROM SDTRX50TB "
            sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND TCRCD  = '" & Trim(DBRec.Fields("TCRCD")) & "'"
            sStr = sStr & "    AND SUBJCD = '" & Trim(DBRec.Fields("SUBJCD")) & "'"
            sStr = sStr & "    AND WEEKS  = " & aWeek
            sStr = sStr & "    AND LESSON = " & aLesson
    
    
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
                            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
                    
            If nExe = 1 Then
                
                '<< 해당강사 삭제 >>
                With sprTmr_Tcr
                    For nRow = 1 To .MaxRows Step 1
                        .Row = nRow:        nChkRow = .Row
                        .Col = SpreadHeader:            sTmpTcr = Trim(.Text)
                        .Col = SpreadHeader + 1:        sTmpSubj = Trim(.Text)
                        
                        If StrComp(Trim(DBRec.Fields("TCRCD")), sTmpTcr, vbTextCompare) = 0 And _
                           StrComp(Trim(DBRec.Fields("SUBJCD")), sTmpSubj, vbTextCompare) = 0 Then
                           
                            For nCol = 1 To .MaxCols Step 1
                                .Col = nCol:        nChkCol = .Col
                                .Row = SpreadHeader + 1:    sTmpWeek = Trim(.Text)
                                .Row = SpreadHeader + 2:    sTmpLesson = Trim(.Text)
                                
                                If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                                   StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                                    
                                    .Row = nChkRow
                                    .Col = nChkCol
                                        .Text = ""
                                    
                                    If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                                       sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                                        ' no action
                                    Else
                                        .Row2 = .Row
                                        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    End If
                                    
                                    Exit For
                                End If
                            Next nCol
                        End If
                    Next nRow
                End With
                
                basDataBase.DBConn.CommitTrans
                
            End If
            
            DBRec.MoveNext
        Next nRec
    End If
            
    Find_TCR_and_Del_LSN = bRet
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    Find_TCR_and_Del_LSN = False

End Function

'## 과목/강사로 처리
Private Function Find_TCR_and_Del_TCR(ByVal aLsnCD As String, ByVal aWeek As String, ByVal aLesson As String) As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim bRet        As Boolean
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTmpTcr     As String
    Dim sTmpSubj    As String
    Dim sTmpWeek    As String
    Dim sTmpLesson  As String
    
    Dim nExe        As Long
    
    Dim nChkRow     As Long
    Dim nChkCol     As Long
    
    
    On Error GoTo ErrStmt
    
    bRet = True
    
    sStr = ""
    sStr = sStr & " SELECT A.TCRCD, A.SUBJCD"
    sStr = sStr & "   From SDTRX50TB A, "
    
    sStr = sStr & "        (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 추가
    sStr = sStr & "           FROM SDLSN01TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         UNION"
    sStr = sStr & "         SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "           FROM SDLSN02TB "
    sStr = sStr & "          WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        ) B"

    sStr = sStr & "  WHERE A.ACID   = B.ACID  "
    sStr = sStr & "    AND A.LSNCD  = B.LSNCD "
    sStr = sStr & "    AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND A.LSNCD  = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "    AND A.WEEKS  = " & aWeek
    sStr = sStr & "    AND A.LESSON = " & aLesson
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    If DBRec.RecordCount > 0 Then
        DBRec.MoveFirst
            
        If DBRec.RecordCount > 0 Then
            basDataBase.DBConn.BeginTrans
        End If
            
        For nRec = 1 To DBRec.RecordCount Step 1
            
            sStr = ""
            sStr = sStr & " DELETE"
            sStr = sStr & "   FROM SDTRX50TB "
            sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND TCRCD  = '" & Trim(DBRec.Fields("TCRCD")) & "'"
            sStr = sStr & "    AND SUBJCD = '" & Trim(DBRec.Fields("SUBJCD")) & "'"
            sStr = sStr & "    AND WEEKS  = " & aWeek
            sStr = sStr & "    AND LESSON = " & aLesson
    
    
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
                            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
                    
            If nExe = 1 Then
                
                '<< 해당강사 삭제 >>
                With sprTmr_Tcr
                    For nRow = 1 To .MaxRows Step 1
                        .Row = nRow:        nChkRow = .Row
                        .Col = SpreadHeader:            sTmpTcr = Trim(.Text)
                        .Col = SpreadHeader + 1:        sTmpSubj = Trim(.Text)
                        
                        If StrComp(Trim(DBRec.Fields("TCRCD")), sTmpTcr, vbTextCompare) = 0 And _
                           StrComp(Trim(DBRec.Fields("SUBJCD")), sTmpSubj, vbTextCompare) = 0 Then
                           
                            For nCol = 1 To .MaxCols Step 1
                                .Col = nCol:        nChkCol = .Col
                                .Row = SpreadHeader + 1:    sTmpWeek = Trim(.Text)
                                .Row = SpreadHeader + 2:    sTmpLesson = Trim(.Text)
                                
                                If StrComp(aWeek, sTmpWeek, vbTextCompare) = 0 And _
                                   StrComp(aLesson, sTmpLesson, vbTextCompare) = 0 Then
                                    
                                    .Row = nChkRow
                                    .Col = nChkCol
                                        .Text = ""
                                    
                                    If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                                       sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                                        ' no action
                                    Else
                                        .Row2 = .Row
                                        .Col2 = .Col
                                        .BlockMode = True
                                            .BackColor = basModule.WhiteColor
                                            .BackColorStyle = BackColorStyleUnderGrid
                                        .BlockMode = False
                                    End If
                                    
                                    Exit For
                                End If
                            Next nCol
                        End If
                    Next nRow
                End With
                
                basDataBase.DBConn.CommitTrans
                
            End If
            
            DBRec.MoveNext
        Next nRec
    End If
            
    Find_TCR_and_Del_TCR = bRet
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    Find_TCR_and_Del_TCR = False

End Function
















'#################################################################################################################################################
' 시간표 과목별로 엑셀작업
'#################################################################################################################################################
Private Sub cmdExcelToGwamok_Click()
    Call Make_Tmr_ExcelFile_to_Gwamok
End Sub

Private Sub Make_Tmr_ExcelFile_to_Gwamok()
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim sComp       As String
    
    Dim sFileName   As String
    Dim sFilePath   As String
    Dim sLogFile    As String
    
    Dim nWeekSrt    As Long
    Dim nColor      As Long
    
    Dim nRet        As Long
    Dim nRow2       As Long
    
    
    Dim sTcrTmp     As String
    Dim sTcrComp    As String
    Dim nChkRow     As Long
    
    Dim sTSisu      As String
    Dim sSSisu      As String
    
    If sprTmr_Tcr.MaxRows = 0 Then Exit Sub
    
    If MsgBox("시간표 엑셀자료 생성하시겠습니까?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrDlg

    If Dir(App.Path & "\TMR_EXCEL", vbDirectory) = "" Then MkDir App.Path & "\TMR_EXCEL"

    'TEXT파일을 생성 처리합니다.
    With dlgExcel
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path & "\TMR_EXCEL"
        .Filter = "DAT FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowSave

        '파일명을 처리합니다.
        If (.fileName) = "" Then Exit Sub
        
        sFileName = Mid$(dlgExcel.FileTitle, 1, InStr(1, dlgExcel.FileTitle, ".", vbTextCompare) - 1)
        sFilePath = Mid$(dlgExcel.fileName, 1, Len(dlgExcel.fileName) - InStrB(1, dlgExcel.fileName, "\", vbTextCompare) - 1)
        sLogFile = sFilePath & sFileName & ".LOG"
        
    End With

    On Error GoTo 0
    On Error GoTo ErrExcel
    
    sprExcel.ColHeadersShow = True
    sprExcel.RowHeadersShow = True
    
    sprExcel.MaxRows = 0
    sprExcel.MaxCols = 0
    
    For nRow = 1 To sprTmr_Tcr.ColHeaderRows Step 1
        sprTmr_Tcr.Row = SpreadHeader + nRow - 1
            '< 데이타 복사 >
            sprExcel.MaxRows = sprExcel.MaxRows + 1
            sprExcel.Row = sprExcel.MaxRows                                         '< header row
        
            sprExcel.MaxCols = sprTmr_Tcr.RowHeaderCols + sprTmr_Tcr.MaxCols        '< 전체 cols
        
            '< Row Header 생성 >
            For nCol = 1 To sprTmr_Tcr.RowHeaderCols Step 1
                sprTmr_Tcr.Col = SpreadHeader + nCol - 1
                    sTmp = sprTmr_Tcr.Text
                    
                    sprExcel.Col = nCol                                                 '< 데이터 넣음
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
                    
                    With sprExcel
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.ShadowColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End With
            Next nCol
            
            '< Data >
            For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
                sprTmr_Tcr.Col = nCol
                    sTmp = Trim(sprTmr_Tcr.Text)
                
                    sprExcel.Col = sprTmr_Tcr.RowHeaderCols + nCol
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
            Next nCol
            
            sprExcel.SetCellBorder sprTmr_Tcr.RowHeaderCols + 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor1, CellBorderStyleSolid
            
            With sprExcel
                .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.ShadowColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
    Next nRow
    
    '< Data 부분 >
    For nRow = 1 To sprTmr_Tcr.MaxRows Step 1
        sprTmr_Tcr.Row = nRow
        sprTmr_Tcr.Col = SpreadHeader:      sTcrComp = Trim(sprTmr_Tcr.Text)
        
        '< 데이타 복사 >
        sprExcel.MaxRows = sprExcel.MaxRows + 1
        sprExcel.Row = sprExcel.MaxRows                                         '< header row
        
        '< Row Header 생성 >
        For nCol = 1 To sprTmr_Tcr.RowHeaderCols Step 1
            sprTmr_Tcr.Col = SpreadHeader + nCol - 1
                sTmp = sprTmr_Tcr.Text
                
                sprExcel.Col = nCol                                                 '< 데이터 넣음
                Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprExcel.ColWidth(sprExcel.Col) = 5
        Next nCol
                
        '< Data >
        For nCol = 1 To sprTmr_Tcr.MaxCols Step 1
            sprTmr_Tcr.Col = nCol
                If sprTmr_Tcr.BackColor = basModule.SectionColor1 Or _
                   sprTmr_Tcr.BackColor = lblNotTeaching.BackColor Then
                    nColor = sprTmr_Tcr.BackColor
                Else
                    nColor = basModule.WhiteColor
                End If
                
                sTmp = Trim(sprTmr_Tcr.Text)
                
                sprExcel.Col = sprTmr_Tcr.RowHeaderCols + nCol
                sprExcel.ColWidth(sprExcel.Col) = 3
                
                If Trim(sTmp) = "" Then
                    ' no action
                    If nColor = lblNotTeaching.BackColor Then
                        If Trim(sprExcel.Text) <> "#" Then
                            sTmp = "#" & Trim(sprExcel.Text)
                            Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        End If
                    End If
                Else
                    If Trim(sprExcel.Text) <> "" Then
                        If StrComp(Trim(sprExcel.Text), "#", vbTextCompare) = 0 Then
                            sTmp = "#" & sTmp
                        Else
                            If nColor = lblNotTeaching.BackColor Then
                                sTmp = "#" & sTmp
                            Else
                                sTmp = sTmp & "/" & Trim(sprExcel.Text)
                            End If
                        End If
                        Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    Else
                        If nColor = lblNotTeaching.BackColor Then
                            sTmp = "#" & sTmp
                        End If
                        Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    End If
                End If
                
                sprExcel.Row2 = sprExcel.Row
                sprExcel.Col2 = sprExcel.Col
                sprExcel.BlockMode = True
                    sprExcel.BackColor = nColor
                    sprExcel.BackColorStyle = BackColorStyleUnderGrid
                sprExcel.BlockMode = False
        Next nCol
            
        If sprExcel.Row > 4 Then
            If ((sprExcel.Row - sprTmr_Tcr.ColHeaderRows) Mod 5) = 0 Then sprExcel.SetCellBorder 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
        End If
        
        With sprExcel
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = sprTmr_Tcr.RowHeaderCols
            .BlockMode = True
                .BackColor = basModule.ShadowColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End With
    
    Next nRow
        
    '< 데이터 정렬 및 맞춤 >
    With sprExcel
        If .MaxRows > 1 Then
            .SetCellBorder 1, 1, .MaxCols, .MaxRows, 16, &H80000008, CellBorderStyleSolid

            .Row = 3
                .SetCellBorder 1, .Row, .MaxCols, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid

            .Row = 2
            nWeekSrt = 0
            For nCol = 1 To .MaxCols Step 1
                .Col = nCol:    sTmp = Trim(.Text):     If sComp = "" Then sComp = Trim(.Text)
                If Trim(.Text) <> "" Then
                    If StrComp(sComp, sTmp, vbTextCompare) <> 0 Then
                        .SetCellBorder .Col, 1, .Col, .MaxRows, 1, basModule.SectionColor2, CellBorderStyleSolid
                        sComp = sTmp

                        If nWeekSrt = 0 Then
                            .AddCellSpan sprTmr_Tcr.RowHeaderCols + 1, 1, nCol - sprTmr_Tcr.RowHeaderCols - 1, 1
                        Else
                            .AddCellSpan nWeekSrt, 1, nCol - nWeekSrt, 1
                        End If
                        nWeekSrt = nCol

                    End If
                End If
            Next nCol
            If nWeekSrt = 0 Then
                .AddCellSpan sprTmr_Tcr.RowHeaderCols + 1, 1, nCol - sprTmr_Tcr.RowHeaderCols - 1, 1
            Else
                .AddCellSpan nWeekSrt, 1, .MaxCols - nWeekSrt + 1, 1
            End If
            nWeekSrt = nCol

            .Row = 2:       .DeleteRows .Row, 1:        .MaxRows = .MaxRows - 1

            .Col = 7
                .SetCellBorder .Col, 1, .Col, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        End If
    End With
    
    nRet = sprExcel.ExportToExcel(dlgExcel.fileName, "Time_Schedule", sLogFile)
    
    MsgBox "엑셀작성하였습니다." & vbCrLf & _
           "확인하십시요.", vbInformation + vbOKOnly, "시간표 엑셀자료 만들기"

    Exit Sub
ErrExcel:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 엑셀자료 만들기"
    Exit Sub
ErrDlg:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성을 취소하였습니다.", vbCritical + vbOKOnly, "엑셀자료 생성"
End Sub














