VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PRT051 
   Caption         =   "시간표 출력 >> 빈 양식지 출력 - CP"
   ClientHeight    =   10695
   ClientLeft      =   7500
   ClientTop       =   3225
   ClientWidth     =   14850
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10695
   ScaleWidth      =   14850
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Height          =   495
      Left            =   30
      TabIndex        =   333
      Top             =   30
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   334
         Top             =   30
         Width           =   14385
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13170
            TabIndex        =   10
            Text            =   "txtPage"
            Top             =   30
            Width           =   735
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "전체페이지 출력"
            Height          =   375
            Left            =   11100
            TabIndex        =   8
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "현재페이지 출력"
            Height          =   375
            Left            =   9540
            TabIndex        =   7
            Top             =   30
            Width           =   1515
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   450
            Style           =   2  '드롭다운 목록
            TabIndex        =   0
            Top             =   67
            Width           =   1155
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "시간표 조회"
            Height          =   375
            Left            =   5280
            TabIndex        =   4
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtLsn 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Index           =   0
            Left            =   1950
            TabIndex        =   1
            Text            =   "txtLsn"
            Top             =   67
            Width           =   1185
         End
         Begin VB.TextBox txtLsn 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3150
            TabIndex        =   2
            Text            =   "txtLsn"
            Top             =   67
            Width           =   615
         End
         Begin VB.CommandButton cmdTime_in 
            Caption         =   "시간 조회"
            Height          =   375
            Left            =   7020
            TabIndex        =   5
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton cmdinFo_in 
            Caption         =   "안내 조회"
            Height          =   375
            Left            =   8130
            TabIndex        =   6
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "◀"
            Height          =   375
            Left            =   12720
            TabIndex        =   9
            Top             =   30
            Width           =   465
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "▶"
            Height          =   375
            Left            =   13920
            TabIndex        =   11
            Top             =   30
            Width           =   465
         End
         Begin EditLib.fpMask fpYM 
            Height          =   285
            Left            =   3960
            TabIndex        =   3
            Top             =   60
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "반"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1740
            TabIndex        =   336
            Top             =   120
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "계열"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   335
            Top             =   120
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox pReportControl 
      BorderStyle     =   0  '없음
      Height          =   9765
      Left            =   30
      ScaleHeight     =   9765
      ScaleWidth      =   14445
      TabIndex        =   12
      Top             =   540
      Width           =   14445
      Begin VB.PictureBox pReportViewer 
         BackColor       =   &H00FFFFFF&
         Height          =   9765
         Left            =   0
         ScaleHeight     =   9705
         ScaleWidth      =   14175
         TabIndex        =   14
         Top             =   0
         Width           =   14235
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   9810
            TabIndex        =   345
            Text            =   "RTB"
            Top             =   5160
            Width           =   3225
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   8700
            TabIndex        =   344
            Text            =   "RTB"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   8700
            TabIndex        =   343
            Text            =   "RTB"
            Top             =   5040
            Width           =   645
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   2580
            TabIndex        =   342
            Text            =   "LTB"
            Top             =   5160
            Width           =   3225
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   1500
            TabIndex        =   341
            Text            =   "LTB"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   1500
            TabIndex        =   340
            Text            =   "LTB"
            Top             =   5040
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   9390
            TabIndex        =   330
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   9390
            TabIndex        =   329
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   91
            Left            =   8700
            TabIndex        =   328
            Text            =   "08:00"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   8700
            TabIndex        =   327
            Text            =   "08:00"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   9390
            TabIndex        =   326
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   9390
            TabIndex        =   325
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   81
            Left            =   8700
            TabIndex        =   324
            Text            =   "08:00"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   8700
            TabIndex        =   323
            Text            =   "08:00"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   9390
            TabIndex        =   322
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   9390
            TabIndex        =   321
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   71
            Left            =   8700
            TabIndex        =   320
            Text            =   "08:00"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   8700
            TabIndex        =   319
            Text            =   "08:00"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   9390
            TabIndex        =   318
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   9390
            TabIndex        =   317
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   61
            Left            =   8700
            TabIndex        =   316
            Text            =   "08:00"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   8700
            TabIndex        =   315
            Text            =   "08:00"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   9390
            TabIndex        =   314
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   9390
            TabIndex        =   313
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   51
            Left            =   8700
            TabIndex        =   312
            Text            =   "08:00"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   8700
            TabIndex        =   311
            Text            =   "08:00"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   9390
            TabIndex        =   310
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   9390
            TabIndex        =   309
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   41
            Left            =   8700
            TabIndex        =   308
            Text            =   "08:00"
            Top             =   2940
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   8700
            TabIndex        =   307
            Text            =   "08:00"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   9390
            TabIndex        =   306
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   9390
            TabIndex        =   305
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   31
            Left            =   8700
            TabIndex        =   304
            Text            =   "08:00"
            Top             =   2520
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   8700
            TabIndex        =   303
            Text            =   "08:00"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   9390
            TabIndex        =   302
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   9390
            TabIndex        =   301
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   8700
            TabIndex        =   300
            Text            =   "08:00"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   9390
            TabIndex        =   299
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   278
            Left            =   9600
            TabIndex        =   298
            Text            =   "월"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   277
            Left            =   10350
            TabIndex        =   297
            Text            =   "화"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   276
            Left            =   11070
            TabIndex        =   296
            Text            =   "수"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   275
            Left            =   11820
            TabIndex        =   295
            Text            =   "목"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   274
            Left            =   12540
            TabIndex        =   294
            Text            =   "금"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   273
            Left            =   13260
            TabIndex        =   293
            Text            =   "토"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   272
            Left            =   7980
            TabIndex        =   292
            Text            =   "1교시"
            Top             =   1770
            Width           =   585
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   9390
            TabIndex        =   291
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   270
            Left            =   7980
            TabIndex        =   290
            Text            =   "2교시"
            Top             =   2190
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   269
            Left            =   7980
            TabIndex        =   289
            Text            =   "3교시"
            Top             =   2610
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   268
            Left            =   7980
            TabIndex        =   288
            Text            =   "4교시"
            Top             =   3000
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   267
            Left            =   7980
            TabIndex        =   287
            Text            =   "5교시"
            Top             =   3840
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   266
            Left            =   7980
            TabIndex        =   286
            Text            =   "6교시"
            Top             =   4290
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   265
            Left            =   7980
            TabIndex        =   285
            Text            =   "7교시"
            Top             =   4680
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   264
            Left            =   7980
            TabIndex        =   284
            Text            =   "8교시"
            Top             =   5520
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   263
            Left            =   7980
            TabIndex        =   283
            Text            =   "9교시"
            Top             =   5970
            Width           =   585
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   8700
            TabIndex        =   282
            Text            =   "08:00"
            Top             =   1680
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   8700
            TabIndex        =   281
            Text            =   "08:00"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   10110
            TabIndex        =   280
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   10110
            TabIndex        =   279
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   10110
            TabIndex        =   278
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   10110
            TabIndex        =   277
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   10110
            TabIndex        =   276
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   10110
            TabIndex        =   275
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   10110
            TabIndex        =   274
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   10110
            TabIndex        =   273
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   10110
            TabIndex        =   272
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   10110
            TabIndex        =   271
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   10110
            TabIndex        =   270
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   10110
            TabIndex        =   269
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   10110
            TabIndex        =   268
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   10110
            TabIndex        =   267
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   10110
            TabIndex        =   266
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   10110
            TabIndex        =   265
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   10110
            TabIndex        =   264
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   10110
            TabIndex        =   263
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   10830
            TabIndex        =   262
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   10830
            TabIndex        =   261
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   10830
            TabIndex        =   260
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   10830
            TabIndex        =   259
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   10830
            TabIndex        =   258
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   10830
            TabIndex        =   257
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   10830
            TabIndex        =   256
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   10830
            TabIndex        =   255
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   10830
            TabIndex        =   254
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   10830
            TabIndex        =   253
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   10830
            TabIndex        =   252
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   10830
            TabIndex        =   251
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   10830
            TabIndex        =   250
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   10830
            TabIndex        =   249
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   10830
            TabIndex        =   248
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   10830
            TabIndex        =   247
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   10830
            TabIndex        =   246
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   10830
            TabIndex        =   245
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   11580
            TabIndex        =   244
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   11580
            TabIndex        =   243
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   11580
            TabIndex        =   242
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   11580
            TabIndex        =   241
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   11580
            TabIndex        =   240
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   11580
            TabIndex        =   239
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   11580
            TabIndex        =   238
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   11580
            TabIndex        =   237
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   11580
            TabIndex        =   236
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   11580
            TabIndex        =   235
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   11580
            TabIndex        =   234
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   11580
            TabIndex        =   233
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   11580
            TabIndex        =   232
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   11580
            TabIndex        =   231
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   11580
            TabIndex        =   230
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   11580
            TabIndex        =   229
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   11580
            TabIndex        =   228
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   11580
            TabIndex        =   227
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   12330
            TabIndex        =   226
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   12330
            TabIndex        =   225
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   12330
            TabIndex        =   224
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   12330
            TabIndex        =   223
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   12330
            TabIndex        =   222
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   12330
            TabIndex        =   221
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   12330
            TabIndex        =   220
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   12330
            TabIndex        =   219
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   12330
            TabIndex        =   218
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   12330
            TabIndex        =   217
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   12330
            TabIndex        =   216
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   12330
            TabIndex        =   215
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   12330
            TabIndex        =   214
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   12330
            TabIndex        =   213
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   12330
            TabIndex        =   212
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   12330
            TabIndex        =   211
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   12330
            TabIndex        =   210
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   12330
            TabIndex        =   209
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   13050
            TabIndex        =   208
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   13050
            TabIndex        =   207
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   13050
            TabIndex        =   206
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   13050
            TabIndex        =   205
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   13050
            TabIndex        =   204
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   13050
            TabIndex        =   203
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   13050
            TabIndex        =   202
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   13050
            TabIndex        =   201
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   13050
            TabIndex        =   200
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   13050
            TabIndex        =   199
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   13050
            TabIndex        =   198
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   13050
            TabIndex        =   197
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   13050
            TabIndex        =   196
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   13050
            TabIndex        =   195
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   13050
            TabIndex        =   194
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   13050
            TabIndex        =   193
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox RT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   13050
            TabIndex        =   192
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox RS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   13050
            TabIndex        =   191
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   5850
            TabIndex        =   190
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   5850
            TabIndex        =   189
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   5850
            TabIndex        =   188
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   5850
            TabIndex        =   187
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   5850
            TabIndex        =   186
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   5850
            TabIndex        =   185
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   5850
            TabIndex        =   184
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   5850
            TabIndex        =   183
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   5850
            TabIndex        =   182
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   5850
            TabIndex        =   181
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   5850
            TabIndex        =   180
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   5850
            TabIndex        =   179
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   5850
            TabIndex        =   178
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   5850
            TabIndex        =   177
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   5850
            TabIndex        =   176
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   5850
            TabIndex        =   175
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   5850
            TabIndex        =   174
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   5850
            TabIndex        =   173
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   5130
            TabIndex        =   172
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   5130
            TabIndex        =   171
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   5130
            TabIndex        =   170
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   5130
            TabIndex        =   169
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   5130
            TabIndex        =   168
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   5130
            TabIndex        =   167
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   5130
            TabIndex        =   166
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   5130
            TabIndex        =   165
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   5130
            TabIndex        =   164
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   5130
            TabIndex        =   163
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   5130
            TabIndex        =   162
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   5130
            TabIndex        =   161
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   5130
            TabIndex        =   160
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   5130
            TabIndex        =   159
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   5130
            TabIndex        =   158
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   5130
            TabIndex        =   157
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   5130
            TabIndex        =   156
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   5130
            TabIndex        =   155
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   4380
            TabIndex        =   154
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   4380
            TabIndex        =   153
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   4380
            TabIndex        =   152
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   4380
            TabIndex        =   151
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   4380
            TabIndex        =   150
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   4380
            TabIndex        =   149
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   4380
            TabIndex        =   148
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   4380
            TabIndex        =   147
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   4380
            TabIndex        =   146
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   4380
            TabIndex        =   145
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   4380
            TabIndex        =   144
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   4380
            TabIndex        =   143
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   4380
            TabIndex        =   142
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   4380
            TabIndex        =   141
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   4380
            TabIndex        =   140
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   4380
            TabIndex        =   139
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   4380
            TabIndex        =   138
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   4380
            TabIndex        =   137
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   3630
            TabIndex        =   136
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   3630
            TabIndex        =   135
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   3630
            TabIndex        =   134
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   3630
            TabIndex        =   133
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   3630
            TabIndex        =   132
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   3630
            TabIndex        =   131
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   3630
            TabIndex        =   130
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   3630
            TabIndex        =   129
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   3630
            TabIndex        =   128
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   3630
            TabIndex        =   127
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   3630
            TabIndex        =   126
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   3630
            TabIndex        =   125
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   3630
            TabIndex        =   124
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   3630
            TabIndex        =   123
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   3630
            TabIndex        =   122
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   3630
            TabIndex        =   121
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   3630
            TabIndex        =   120
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   3630
            TabIndex        =   119
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   2910
            TabIndex        =   118
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   2910
            TabIndex        =   117
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   2910
            TabIndex        =   116
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   2910
            TabIndex        =   115
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   2910
            TabIndex        =   114
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   2910
            TabIndex        =   113
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   2910
            TabIndex        =   112
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   2910
            TabIndex        =   111
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   2910
            TabIndex        =   110
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   2910
            TabIndex        =   109
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   2910
            TabIndex        =   108
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   2910
            TabIndex        =   107
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   2910
            TabIndex        =   106
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   2910
            TabIndex        =   105
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   2910
            TabIndex        =   104
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   2910
            TabIndex        =   103
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   2910
            TabIndex        =   102
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   2910
            TabIndex        =   101
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   1500
            TabIndex        =   100
            Text            =   "08:00"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   1500
            TabIndex        =   99
            Text            =   "08:00"
            Top             =   1680
            Width           =   645
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   780
            TabIndex        =   98
            Text            =   "9교시"
            Top             =   5970
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   780
            TabIndex        =   97
            Text            =   "8교시"
            Top             =   5520
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   780
            TabIndex        =   96
            Text            =   "7교시"
            Top             =   4680
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   780
            TabIndex        =   95
            Text            =   "6교시"
            Top             =   4290
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   780
            TabIndex        =   94
            Text            =   "5교시"
            Top             =   3840
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   780
            TabIndex        =   93
            Text            =   "4교시"
            Top             =   3000
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   780
            TabIndex        =   92
            Text            =   "3교시"
            Top             =   2610
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   780
            TabIndex        =   91
            Text            =   "2교시"
            Top             =   2190
            Width           =   585
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   2190
            TabIndex        =   90
            Text            =   "언A"
            Top             =   1650
            Width           =   645
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   780
            TabIndex        =   89
            Text            =   "1교시"
            Top             =   1770
            Width           =   585
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   6060
            TabIndex        =   88
            Text            =   "토"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   5340
            TabIndex        =   87
            Text            =   "금"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   4620
            TabIndex        =   86
            Text            =   "목"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3870
            TabIndex        =   85
            Text            =   "수"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3150
            TabIndex        =   84
            Text            =   "화"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox Labels 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   2400
            TabIndex        =   83
            Text            =   "월"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   2190
            TabIndex        =   82
            Text            =   "유하균"
            Top             =   1860
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   1500
            TabIndex        =   81
            Text            =   "08:00"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   1500
            TabIndex        =   80
            Text            =   "08:00"
            Top             =   2100
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   2190
            TabIndex        =   79
            Text            =   "언A"
            Top             =   2070
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   2190
            TabIndex        =   78
            Text            =   "유하균"
            Top             =   2280
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   1500
            TabIndex        =   77
            Text            =   "08:00"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   31
            Left            =   1500
            TabIndex        =   76
            Text            =   "08:00"
            Top             =   2520
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   2190
            TabIndex        =   75
            Text            =   "언A"
            Top             =   2490
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   2190
            TabIndex        =   74
            Text            =   "유하균"
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   1500
            TabIndex        =   73
            Text            =   "08:00"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   41
            Left            =   1500
            TabIndex        =   72
            Text            =   "08:00"
            Top             =   2940
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   2190
            TabIndex        =   71
            Text            =   "언A"
            Top             =   2910
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   2190
            TabIndex        =   70
            Text            =   "유하균"
            Top             =   3120
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   1500
            TabIndex        =   69
            Text            =   "08:00"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   51
            Left            =   1500
            TabIndex        =   68
            Text            =   "08:00"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   2190
            TabIndex        =   67
            Text            =   "언A"
            Top             =   3750
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   2190
            TabIndex        =   66
            Text            =   "유하균"
            Top             =   3960
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   1500
            TabIndex        =   65
            Text            =   "08:00"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   61
            Left            =   1500
            TabIndex        =   64
            Text            =   "08:00"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   2190
            TabIndex        =   63
            Text            =   "언A"
            Top             =   4170
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   2190
            TabIndex        =   62
            Text            =   "유하균"
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   1500
            TabIndex        =   61
            Text            =   "08:00"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   71
            Left            =   1500
            TabIndex        =   60
            Text            =   "08:00"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   2190
            TabIndex        =   59
            Text            =   "언A"
            Top             =   4620
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   2190
            TabIndex        =   58
            Text            =   "유하균"
            Top             =   4830
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   1500
            TabIndex        =   57
            Text            =   "08:00"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   81
            Left            =   1500
            TabIndex        =   56
            Text            =   "08:00"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   2190
            TabIndex        =   55
            Text            =   "언A"
            Top             =   5460
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   2190
            TabIndex        =   54
            Text            =   "유하균"
            Top             =   5670
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   1500
            TabIndex        =   53
            Text            =   "08:00"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox LC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   91
            Left            =   1500
            TabIndex        =   52
            Text            =   "08:00"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LS 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   2190
            TabIndex        =   51
            Text            =   "언A"
            Top             =   5880
            Width           =   645
         End
         Begin VB.TextBox LT 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   2190
            TabIndex        =   50
            Text            =   "유하균"
            Top             =   6090
            Width           =   645
         End
         Begin VB.TextBox RC 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   8700
            TabIndex        =   49
            Text            =   "08:00"
            Top             =   2100
            Width           =   645
         End
         Begin VB.TextBox LHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   750
            TabIndex        =   48
            Text            =   "계열 : 인문계"
            Top             =   1020
            Width           =   1245
         End
         Begin VB.TextBox LHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   2160
            TabIndex        =   47
            Text            =   "반 : 언어영역반"
            Top             =   1020
            Width           =   1395
         End
         Begin VB.TextBox LHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   3750
            TabIndex        =   46
            Text            =   "교실 : 100 호"
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox LHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   5100
            TabIndex        =   45
            Text            =   "담담 : 유하균"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox RHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   7950
            TabIndex        =   44
            Text            =   "계열 : 인문계"
            Top             =   1020
            Width           =   1245
         End
         Begin VB.TextBox RHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   9330
            TabIndex        =   43
            Text            =   "반 : 언어영역반"
            Top             =   1020
            Width           =   1395
         End
         Begin VB.TextBox RHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   10950
            TabIndex        =   42
            Text            =   "교실 : 100 호"
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox RHD 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   12330
            TabIndex        =   41
            Text            =   "담당 : 유하균"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   40
            Text            =   "ML"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   1
            Left            =   720
            TabIndex        =   39
            Text            =   "ML"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   2
            Left            =   720
            TabIndex        =   38
            Text            =   "ML"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   37
            Text            =   "ML"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   4
            Left            =   720
            TabIndex        =   36
            Text            =   "ML"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   5
            Left            =   720
            TabIndex        =   35
            Text            =   "ML"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   6
            Left            =   720
            TabIndex        =   34
            Text            =   "ML"
            Top             =   8250
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   0
            Left            =   7920
            TabIndex        =   33
            Text            =   "MR"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   1
            Left            =   7920
            TabIndex        =   32
            Text            =   "MR"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   2
            Left            =   7920
            TabIndex        =   31
            Text            =   "MR"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   3
            Left            =   7920
            TabIndex        =   30
            Text            =   "MR"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   4
            Left            =   7920
            TabIndex        =   29
            Text            =   "MR"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   5
            Left            =   7920
            TabIndex        =   28
            Text            =   "MR"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   6
            Left            =   7920
            TabIndex        =   27
            Text            =   "MR"
            Top             =   8250
            Width           =   5895
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   1500
            TabIndex        =   26
            Text            =   "LTB"
            Top             =   3330
            Width           =   645
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   1500
            TabIndex        =   25
            Text            =   "LTB"
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   2580
            TabIndex        =   24
            Text            =   "LTB"
            Top             =   3450
            Width           =   3225
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   8700
            TabIndex        =   23
            Text            =   "RTB"
            Top             =   3330
            Width           =   645
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   8700
            TabIndex        =   22
            Text            =   "RTB"
            Top             =   3540
            Width           =   615
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   9810
            TabIndex        =   21
            Text            =   "RTB"
            Top             =   3450
            Width           =   3225
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   7
            Left            =   720
            TabIndex        =   20
            Text            =   "ML"
            Top             =   8520
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   8
            Left            =   720
            TabIndex        =   19
            Text            =   "ML"
            Top             =   8790
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   9
            Left            =   720
            TabIndex        =   18
            Text            =   "ML"
            Top             =   9060
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   7
            Left            =   7920
            TabIndex        =   17
            Text            =   "MR"
            Top             =   8520
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   8
            Left            =   7920
            TabIndex        =   16
            Text            =   "MR"
            Top             =   8790
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   270
            Index           =   9
            Left            =   7920
            TabIndex        =   15
            Text            =   "MR"
            Top             =   9060
            Width           =   5895
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   60
            X1              =   8610
            X2              =   8610
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   59
            X1              =   10080
            X2              =   10080
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   58
            X1              =   10800
            X2              =   10800
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   57
            X1              =   11550
            X2              =   11550
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   56
            X1              =   12300
            X2              =   12300
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   55
            X1              =   13020
            X2              =   13020
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   54
            X1              =   13020
            X2              =   13020
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   53
            X1              =   8610
            X2              =   8610
            Y1              =   1620
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   52
            X1              =   12300
            X2              =   12300
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   51
            X1              =   11550
            X2              =   11550
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   49
            X1              =   10800
            X2              =   10800
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   47
            X1              =   10080
            X2              =   10080
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   45
            X1              =   1410
            X2              =   1410
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   2880
            X2              =   2880
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   44
            X1              =   3600
            X2              =   3600
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   43
            X1              =   4350
            X2              =   4350
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   5100
            X2              =   5100
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   41
            X1              =   5820
            X2              =   5820
            Y1              =   3720
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   3
            X1              =   5820
            X2              =   5820
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   1410
            X2              =   1410
            Y1              =   1620
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   5100
            X2              =   5100
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   4350
            X2              =   4350
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   6
            X1              =   3600
            X2              =   3600
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   2880
            X2              =   2880
            Y1              =   1260
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   720
            X2              =   6570
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   50
            X1              =   9360
            X2              =   9360
            Y1              =   1260
            Y2              =   6300
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   1
            Left            =   7920
            Top             =   1260
            Width           =   5865
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   48
            X1              =   7920
            X2              =   13770
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   40
            X1              =   7920
            X2              =   13770
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   39
            X1              =   7920
            X2              =   13770
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   38
            X1              =   7920
            X2              =   13770
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   37
            X1              =   7920
            X2              =   13770
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   36
            X1              =   7920
            X2              =   13770
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   35
            X1              =   7920
            X2              =   13770
            Y1              =   4140
            Y2              =   4140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   34
            X1              =   7920
            X2              =   13770
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   33
            X1              =   7920
            X2              =   13770
            Y1              =   5010
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   32
            X1              =   7920
            X2              =   13770
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   31
            X1              =   7920
            X2              =   13770
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   30
            X1              =   8610
            X2              =   8610
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   29
            X1              =   10080
            X2              =   10080
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   19
            X1              =   10800
            X2              =   10800
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   18
            X1              =   11550
            X2              =   11550
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   13
            X1              =   12300
            X2              =   12300
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   12
            X1              =   13020
            X2              =   13020
            Y1              =   5430
            Y2              =   6270
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   27
            X1              =   5820
            X2              =   5820
            Y1              =   5430
            Y2              =   6270
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   26
            X1              =   5100
            X2              =   5100
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   25
            X1              =   4350
            X2              =   4350
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   24
            X1              =   3600
            X2              =   3600
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   23
            X1              =   2880
            X2              =   2880
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   22
            X1              =   1410
            X2              =   1410
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   8
            X1              =   720
            X2              =   6570
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   21
            X1              =   720
            X2              =   6570
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   20
            X1              =   720
            X2              =   6570
            Y1              =   5010
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   17
            X1              =   720
            X2              =   6570
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   720
            X2              =   6570
            Y1              =   4140
            Y2              =   4140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   11
            X1              =   720
            X2              =   6570
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   10
            X1              =   720
            X2              =   6570
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   9
            X1              =   720
            X2              =   6570
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   7
            X1              =   720
            X2              =   6570
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   0
            X1              =   2160
            X2              =   2160
            Y1              =   1260
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   28
            X1              =   720
            X2              =   6570
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   0
            Left            =   720
            Top             =   1260
            Width           =   5865
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '점
            Index           =   15
            X1              =   7260
            X2              =   7260
            Y1              =   90
            Y2              =   9660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "시   간   표"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2340
            TabIndex        =   332
            Top             =   300
            Width           =   2205
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "시   간   표"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9570
            TabIndex        =   331
            Top             =   300
            Width           =   2205
         End
         Begin VB.Shape FillBOXs2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   555
            Index           =   2
            Left            =   1440
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   4035
         End
         Begin VB.Shape FillBOXs2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   555
            Index           =   0
            Left            =   8640
            Shape           =   4  '둥근 사각형
            Top             =   210
            Width           =   4035
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9765
         Left            =   14220
         Max             =   1
         TabIndex        =   13
         Top             =   0
         Width           =   225
      End
   End
   Begin FPSpread.vaSpread sprLsn 
      Height          =   6255
      Left            =   2790
      TabIndex        =   337
      Top             =   10680
      Width           =   2685
      _Version        =   393216
      _ExtentX        =   4736
      _ExtentY        =   11033
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
      ScrollBars      =   2
      SpreadDesigner  =   "PRT051.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   14640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread sprinFo 
      Height          =   4395
      Left            =   15090
      TabIndex        =   338
      Top             =   6630
      Width           =   6045
      _Version        =   393216
      _ExtentX        =   10663
      _ExtentY        =   7752
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
      MaxCols         =   1
      MaxRows         =   12
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "PRT051.frx":02B4
   End
   Begin FPSpread.vaSpread sprTime 
      Height          =   5535
      Left            =   15090
      TabIndex        =   339
      Top             =   840
      Width           =   1425
      _Version        =   393216
      _ExtentX        =   2514
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
      MaxCols         =   1
      MaxRows         =   22
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "PRT051.frx":0704
   End
End
Attribute VB_Name = "PRT051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : PRT051
'   모 듈  목 적 : 반별 시간표 출력
'
'   작   성   일 : 2008/02/19
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Type tTimeTable
    '<< 비교 KEY VALUE >>
    LSNCD           As String
    
    '< DATA >
    GAEYUL          As String
    LSNNM           As String
    CLASS_NM        As String
    DAMIM           As String
    
    DATA(110, 2)    As String
End Type
Private uTimeTable()    As tTimeTable


Private sini_Path   As String


Private Sub Form_Click()
    sprLsn.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
End Sub

Private Sub Form_Load()

    Dim nRow        As Long

    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    basFunction.RemoveContextMenu txtLsn(0)
    
    fpYM.Text = Format(Now, "YYYYMM")
    
    Me.Tag = "LOAD"
        
        Me.Width = 14600
        Me.Height = 10755
        
        sini_Path = App.Path & "\DAESUNG.INI"       '<< ini file
        cmdTime_in.Caption = "시간 조회"
        cmdinFo_in.Caption = "안내 조회"
        
        '>> sprTime
        cmdTime_in.Tag = ""
        With sprTime
            .Top = 480
            .Left = 6510
        
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                    .Text = ""
                    
                If (nRow Mod 2) = 0 Then
                    Call .SetCellBorder(.Col, .Row, .Col, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid)
                End If
                
            Next nRow
            
            .ZOrder 0
            .Visible = False
        End With
        
        '>> sprinFo
        cmdinFo_in.Tag = ""
        With sprinFo
            .Top = 480
            .Left = 8100
        
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                    .Text = ""
                    
                Call .SetCellBorder(.Col, .Row, .Col, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid)
            Next nRow
            
            .ZOrder 0
            .Visible = False
        End With
        
        
        txtLsn(0).Text = ""
        txtLsn(1).Text = ""
        
        txtLsn(0).Tag = ""
        With sprLsn
            .Top = 480
            .Left = 2520
        
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .ZOrder 0
            .MaxRows = 0
            .Visible = False
        End With
        
        
        '>> 계열
        With cboKaeyol
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            .ListIndex = 0
        End With
        
        VScroll1.Min = 1
        VScroll1.Max = 100
        VScroll1.SmallChange = 1
        VScroll1.LargeChange = 1
        VScroll1.Enabled = False
        
        ReDim uTimeTable(0) As tTimeTable
        
        txtPage.Text = ""
        
        Call Clear_Form_Control                 '< CONTROL 초기화
        'Call Test_Print                     '< TEST

        Call init_Display_Time_and_inFo         '< 시간 및 안내내역 => 시간표로
        
        
    Me.Tag = ""
    
End Sub

'## 테스트 출력
Private Sub Test_Print()

    Dim nRow        As Integer
    Dim nCol        As Integer
    
    Dim sinDex      As String
    
    On Error Resume Next
    
    For nRow = 1 To 10 Step 1
        '< 시간 >
        For nCol = 1 To 2 Step 1
            sinDex = Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            LC(CInt(sinDex)).Text = "LC" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            RC(CInt(sinDex)).Text = "RC" & Trim(CStr(nRow)) & Trim(CStr(nCol))
        Next nCol
        
        '< 과목/ 강사내역 test >
        For nCol = 2 To 7 Step 1
            sinDex = Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            LS(CInt(sinDex)).Text = "LS" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            LT(CInt(sinDex)).Text = "LT" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            RS(CInt(sinDex)).Text = "RS" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            RT(CInt(sinDex)).Text = "RT" & Trim(CStr(nRow)) & Trim(CStr(nCol))
        Next nCol
    Next nRow

End Sub


'## control 초기화
Private Sub Clear_Form_Control()
    Dim UsrCtl      As Control
    
    '>> 초기화
    For Each UsrCtl In Me
        With UsrCtl
            If UCase(TypeName(UsrCtl)) = "TEXTBOX" And UCase(UsrCtl.Name) <> "TXTLSN" And UCase(UsrCtl.Name) = "NONPRINTLBL" Then .Text = ""
            If UCase(UsrCtl.Name) = "LC" Or _
               UCase(UsrCtl.Name) = "LS" Or _
               UCase(UsrCtl.Name) = "LT" Or _
               UCase(UsrCtl.Name) = "RC" Or _
               UCase(UsrCtl.Name) = "RS" Or _
               UCase(UsrCtl.Name) = "RT" Or _
               UCase(UsrCtl.Name) = "LHD" Or _
               UCase(UsrCtl.Name) = "RHD" Then
                .Text = ""
            End If
            
            If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
            If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
        End With
    Next
End Sub


'## 시간 및 안내내역 => 시간표로
Private Sub init_Display_Time_and_inFo()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    '## 시간내역
    sGbn = "TIME"
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "11", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(11).Text = sTmp:  RC(11).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "12", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(12).Text = sTmp:  RC(12).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "21", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(21).Text = sTmp:  RC(21).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "22", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(22).Text = sTmp:  RC(22).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "31", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(31).Text = sTmp:  RC(31).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "32", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(32).Text = sTmp:  RC(32).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "41", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(41).Text = sTmp:  RC(41).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "42", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(42).Text = sTmp:  RC(42).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "51", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(51).Text = sTmp:  RC(51).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "52", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(52).Text = sTmp:  RC(52).Text = sTmp
            
        
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(1).Text = sTmp:     RTB(1).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(2).Text = sTmp:     RTB(2).Text = sTmp
            
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "61", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(61).Text = sTmp:  RC(61).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "62", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(62).Text = sTmp:  RC(62).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "71", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(71).Text = sTmp:  RC(71).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "72", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(72).Text = sTmp:  RC(72).Text = sTmp
                                                                                                                                                                      
        '>> 2008.02.25 : 추가
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(3).Text = sTmp:     RTB(3).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(4).Text = sTmp:     RTB(4).Text = sTmp
            
        
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "81", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(81).Text = sTmp:  RC(81).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "82", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(82).Text = sTmp:  RC(82).Text = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "91", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(91).Text = sTmp:  RC(91).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "92", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(92).Text = sTmp:  RC(92).Text = sTmp
        
'        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "101", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'            LC(101).Text = sTmp:  RC(101).Text = sTmp
'        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "102", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'            LC(102).Text = sTmp:  RC(102).Text = sTmp
                        
    
    '## 안내내역
    sGbn = "INFO"
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB1", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(0).Text = sTmp:     RTB(0).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB2", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(5).Text = sTmp:     RTB(5).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(0).Text = sTmp:     MR(0).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(1).Text = sTmp:     MR(1).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(2).Text = sTmp:     MR(2).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(3).Text = sTmp:     MR(3).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO5", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(4).Text = sTmp:     MR(4).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO6", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(5).Text = sTmp:     MR(5).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO7", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(6).Text = sTmp:     MR(6).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO8", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(7).Text = sTmp:     MR(7).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO9", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(8).Text = sTmp:     MR(8).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INF10", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(9).Text = sTmp:     MR(9).Text = sTmp
    
End Sub








'## 시간표 시간 등록 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdTime_in_Click()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    If cmdTime_in.Tag = "" Then
        cmdTime_in.Caption = "시간 등록"
        
        '## 데이터 불러오기
        sprTime.Col = 1
        sGbn = "TIME"
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "11", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(11).Text = sTmp:  RC(11).Text = sTmp:      sprTime.Row = 1:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "12", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(12).Text = sTmp:  RC(12).Text = sTmp:      sprTime.Row = 2:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "21", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(21).Text = sTmp:  RC(21).Text = sTmp:      sprTime.Row = 3:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "22", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(22).Text = sTmp:  RC(22).Text = sTmp:      sprTime.Row = 4:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "31", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(31).Text = sTmp:  RC(31).Text = sTmp:      sprTime.Row = 5:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "32", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(32).Text = sTmp:  RC(32).Text = sTmp:      sprTime.Row = 6:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "41", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(41).Text = sTmp:  RC(41).Text = sTmp:      sprTime.Row = 7:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "42", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(42).Text = sTmp:  RC(42).Text = sTmp:      sprTime.Row = 8:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "51", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(51).Text = sTmp:  RC(51).Text = sTmp:      sprTime.Row = 9:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "52", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(52).Text = sTmp:  RC(52).Text = sTmp:      sprTime.Row = 10:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(1).Text = sTmp:     RTB(1).Text = sTmp:      sprTime.Row = 11:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(2).Text = sTmp:     RTB(2).Text = sTmp:      sprTime.Row = 12:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "61", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(61).Text = sTmp:  RC(61).Text = sTmp:      sprTime.Row = 13:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "62", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(62).Text = sTmp:  RC(62).Text = sTmp:      sprTime.Row = 14:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "71", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(71).Text = sTmp:  RC(71).Text = sTmp:      sprTime.Row = 15:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "72", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(72).Text = sTmp:  RC(72).Text = sTmp:      sprTime.Row = 16:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            
            '>> 2008.02.25 : 추가
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(3).Text = sTmp:     RTB(3).Text = sTmp:      sprTime.Row = 17:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(4).Text = sTmp:     RTB(4).Text = sTmp:      sprTime.Row = 18:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "81", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(81).Text = sTmp:  RC(81).Text = sTmp:      sprTime.Row = 19:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "82", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(82).Text = sTmp:  RC(82).Text = sTmp:      sprTime.Row = 20:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "91", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(91).Text = sTmp:  RC(91).Text = sTmp:      sprTime.Row = 21:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "92", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(92).Text = sTmp:  RC(92).Text = sTmp:      sprTime.Row = 22:        sprTime.value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            
'            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "101", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'                LC(101).Text = sTmp:  RC(101).Text = sTmp:      sprTime.Row = 21:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
'            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "102", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'                LC(102).Text = sTmp:  RC(102).Text = sTmp:      sprTime.Row = 22:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                            
        sprTime.Visible = True
        cmdTime_in.Tag = "SAVE"
        
        sprTime.SetActiveCell 1, 1
        
        Exit Sub
    End If
    
    If MsgBox("시간을 등록하시겠습니까?", vbQuestion + vbYesNo, "시간표 시간등록") = vbNo Then
        cmdTime_in.Caption = "시간 조회"
        sprTime.Visible = False
        cmdTime_in.Tag = ""
        Exit Sub
    End If
    
    If cmdTime_in.Tag = "SAVE" Then
        With sprTime
            sGbn = "TIME"
            
            .Col = 1
            '< 1교시
                .Row = 1:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "11", sTmp, sini_Path): LC(11).Text = sTmp:   RC(11).Text = sTmp
                .Row = 2:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "12", sTmp, sini_Path): LC(12).Text = sTmp:   RC(12).Text = sTmp
            '< 2교시
                .Row = 3:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "21", sTmp, sini_Path): LC(21).Text = sTmp:   RC(21).Text = sTmp
                .Row = 4:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "22", sTmp, sini_Path): LC(22).Text = sTmp:   RC(22).Text = sTmp
            '< 3교시
                .Row = 5:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "31", sTmp, sini_Path): LC(31).Text = sTmp:   RC(31).Text = sTmp
                .Row = 6:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "32", sTmp, sini_Path): LC(32).Text = sTmp:   RC(32).Text = sTmp
            '< 4교시
                .Row = 7:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "41", sTmp, sini_Path): LC(41).Text = sTmp:   RC(41).Text = sTmp
                .Row = 8:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "42", sTmp, sini_Path): LC(42).Text = sTmp:   RC(42).Text = sTmp
            '< 5교시
                .Row = 9:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "51", sTmp, sini_Path): LC(51).Text = sTmp:   RC(51).Text = sTmp
                .Row = 10:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "52", sTmp, sini_Path): LC(52).Text = sTmp:   RC(52).Text = sTmp
                                                                                                                                                     
            '< break
                .Row = 11:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B1", sTmp, sini_Path): LTB(1).Text = sTmp:      RTB(1).Text = sTmp
                .Row = 12:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B2", sTmp, sini_Path): LTB(2).Text = sTmp:      RTB(2).Text = sTmp
                                                                                                                                                     
            '< 6교시
                .Row = 13:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "61", sTmp, sini_Path): LC(61).Text = sTmp:   RC(61).Text = sTmp
                .Row = 14:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "62", sTmp, sini_Path): LC(62).Text = sTmp:   RC(62).Text = sTmp
            '< 7교시
                .Row = 15:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "71", sTmp, sini_Path): LC(71).Text = sTmp:   RC(71).Text = sTmp
                .Row = 16:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "72", sTmp, sini_Path): LC(72).Text = sTmp:   RC(72).Text = sTmp
                    
                    
            '< break : 2008.02.25
                .Row = 17:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B3", sTmp, sini_Path): LTB(3).Text = sTmp:      RTB(3).Text = sTmp
                .Row = 18:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B4", sTmp, sini_Path): LTB(4).Text = sTmp:      RTB(4).Text = sTmp
                    
                    
            '< 8교시
                .Row = 19:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "81", sTmp, sini_Path): LC(81).Text = sTmp:   RC(81).Text = sTmp
                .Row = 20:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "82", sTmp, sini_Path): LC(82).Text = sTmp:   RC(82).Text = sTmp
            '< 9교시
                .Row = 21:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "91", sTmp, sini_Path): LC(91).Text = sTmp:   RC(91).Text = sTmp
                .Row = 22:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "92", sTmp, sini_Path): LC(92).Text = sTmp:   RC(92).Text = sTmp

'            '< 10교시
'                .Row = 21:  sTmp = Left(Trim(.Text), 5)
'                    nRtn = basModule.WritePrivateProfileString(sGbn, "101", sTmp, sini_Path): LC(101).Text = sTmp: RC(101).Text = sTmp
'                .Row = 22:  sTmp = Left(Trim(.Text), 5)
'                    nRtn = basModule.WritePrivateProfileString(sGbn, "102", sTmp, sini_Path): LC(102).Text = sTmp: RC(102).Text = sTmp

        End With
        
        cmdTime_in.Tag = ""
        cmdTime_in.Caption = "시간 조회"
        sprTime.Visible = False
    End If
    
End Sub










Private Sub pReportViewer_Click()
    sprLsn.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
    
End Sub

Private Sub sprTime_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprTime
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = 1
                    .Text = ""
        End Select
    End With
End Sub




'## 시간표 안내등록  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdinFo_in_Click()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    If cmdinFo_in.Tag = "" Then
        cmdinFo_in.Caption = "내용 등록"
        
        '## 데이터 불러오기
        sprinFo.Col = 1
        sGbn = "INFO"
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB1", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(0).Text = sTmp:     RTB(0).Text = sTmp:     sprinFo.Row = 1:        sprinFo.Text = Trim(sTmp)
                
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB2", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(5).Text = sTmp:     RTB(5).Text = sTmp:     sprinFo.Row = 2:        sprinFo.Text = Trim(sTmp)
                
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(0).Text = sTmp:     MR(0).Text = sTmp:     sprinFo.Row = 3:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(1).Text = sTmp:     MR(1).Text = sTmp:     sprinFo.Row = 4:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(2).Text = sTmp:     MR(2).Text = sTmp:     sprinFo.Row = 5:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(3).Text = sTmp:     MR(3).Text = sTmp:     sprinFo.Row = 6:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO5", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(4).Text = sTmp:     MR(4).Text = sTmp:     sprinFo.Row = 7:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO6", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(5).Text = sTmp:     MR(5).Text = sTmp:     sprinFo.Row = 8:          sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO7", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(6).Text = sTmp:     MR(6).Text = sTmp:     sprinFo.Row = 9:          sprinFo.Text = Trim(sTmp)
                
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO8", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(7).Text = sTmp:     MR(7).Text = sTmp:     sprinFo.Row = 10:         sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO9", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(8).Text = sTmp:     MR(8).Text = sTmp:     sprinFo.Row = 11:         sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INF10", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(9).Text = sTmp:     MR(9).Text = sTmp:     sprinFo.Row = 12:         sprinFo.Text = Trim(sTmp)
            
        sprinFo.Visible = True
        cmdinFo_in.Tag = "SAVE"
        
        sprinFo.SetActiveCell 1, 1
        
        Exit Sub
    End If
    
    If MsgBox("안내를 등록하시겠습니까?", vbQuestion + vbYesNo, "시간표 안내등록") = vbNo Then
        cmdinFo_in.Caption = "안내 조회"
        sprinFo.Visible = False
        cmdinFo_in.Tag = ""
        Exit Sub
    End If
    
    If cmdinFo_in.Tag = "SAVE" Then
        With sprinFo
            sGbn = "INFO"
            
            .Col = 1
            '< BREAK
                .Row = 1:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "LRTB1", sTmp, sini_Path):  LTB(0).Text = sTmp: RTB(0).Text = sTmp
                
                .Row = 2:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "LRTB2", sTmp, sini_Path):  LTB(5).Text = sTmp: RTB(5).Text = sTmp
                
                .Row = 3:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO1", sTmp, sini_Path): ML(0).Text = sTmp:  MR(0).Text = sTmp
                .Row = 4:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO2", sTmp, sini_Path): ML(1).Text = sTmp:  MR(1).Text = sTmp
                .Row = 5:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO3", sTmp, sini_Path): ML(2).Text = sTmp:  MR(2).Text = sTmp
                .Row = 6:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO4", sTmp, sini_Path): ML(3).Text = sTmp:  MR(3).Text = sTmp
                .Row = 7:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO5", sTmp, sini_Path): ML(4).Text = sTmp:  MR(4).Text = sTmp
                .Row = 8:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO6", sTmp, sini_Path): ML(5).Text = sTmp:  MR(5).Text = sTmp
                .Row = 9:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO7", sTmp, sini_Path): ML(6).Text = sTmp:  MR(6).Text = sTmp
                    
                .Row = 10:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO8", sTmp, sini_Path): ML(7).Text = sTmp:  MR(7).Text = sTmp
                .Row = 11:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO9", sTmp, sini_Path): ML(8).Text = sTmp:  MR(8).Text = sTmp
                .Row = 12:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INF10", sTmp, sini_Path): ML(9).Text = sTmp:  MR(9).Text = sTmp

        End With
        
        cmdinFo_in.Tag = ""
        cmdinFo_in.Caption = "안내 조회"
        sprinFo.Visible = False
    End If
    
End Sub

Private Sub sprinFo_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprinFo
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = 1
                    .Text = ""
        End Select
    End With
End Sub


'#############################################################################################################################################################




'>> 시간표 조회
Private Sub cmdFind_Click()
    
    On Error GoTo ErrStmt
    
    ReDim uTimeTable(0) As tTimeTable
    
    cmdFind.Enabled = False
        Call Get_TimeTable_Data
        Call Disp_TimeTable_All_Data(1)
        
    cmdFind.Enabled = True
    
    MsgBox "시간표 조회하였습니다.", vbInformation + vbOKOnly, "시간표 조회"
    
    Exit Sub
ErrStmt:
    MsgBox "시간표 조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 조회"
    On Error GoTo 0

End Sub

Private Sub Get_TimeTable_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Long
    Dim nRec        As Long
    Dim sTmp        As String
    
    Dim ninDex      As Long
    
    Dim sLsnCD      As String
    Dim nArray      As Long
    
    On Error GoTo ErrStmt
    
    '>> 초기화 -------------------------------------------------------------------
    Call Clear_Form_Control                 '< CONTROL 초기화
    Call init_Display_Time_and_inFo         '< 시간 및 안내내역 => 시간표로
    '-----------------------------------------------------------------------------
    
    sStr = ""
    
    sStr = sStr & " SELECT LSNCD, LSNNM, KAEYOL, GAEYUL, CLASSNM, DAMIM, IDX, LSNCDNM, TCRNM, SUBJNM"
    sStr = sStr & "   FROM ("
'/* 이동반 ## */
    sStr = sStr & "        SELECT B.LSNCD, B.LSNNM, A.KAEYOL,"
    sStr = sStr & "               DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS GAEYUL,"
    sStr = sStr & "               B.CLASSNM, B.DAMIM,"
    sStr = sStr & "               A.IDX,"
    sStr = sStr & "               A.LSNCDNM,"
    sStr = sStr & "               DECODE(A.LSNNM,'방송수업','','공통') AS TCRNM,"
    sStr = sStr & "               A.LSNNM AS SUBJNM"
    sStr = sStr & "          FROM (SELECT A.ACID, A.LSNNM, NUM AS LSNCDNM, A.KAEYOL, B.WEEKS, B.LESSON, TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX"
    sStr = sStr & "                  FROM (SELECT ACID, TRXCD, LSNNM,"
    sStr = sStr & "                               KAEYOL, B.NUM"
    sStr = sStr & "                          FROM (SELECT ACID, TRXCD, TRXNM AS LSNNM,"
    sStr = sStr & "                                       KAEYOL"
    sStr = sStr & "                                  FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),1,2) AS CUTA,"
    sStr = sStr & "                                               NVL(SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),4,2),'AA') AS CUTB"
    sStr = sStr & "                                          FROM SDTRX01TB"
    sStr = sStr & "                                         WHERE ACID = '" & basModule.schcd & "'"
    sStr = sStr & "                                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                                       )"
    sStr = sStr & "                                 WHERE LTRIM(CUTA,'0123456789') IS NOT NULL"
    sStr = sStr & "                                   AND LTRIM(CUTB,'0123456789') IS NOT NULL"
    sStr = sStr & "                                 ) A,"
    sStr = sStr & "                                SDTRX90TB B"
    sStr = sStr & "                          WHERE B.NO < 40"
    sStr = sStr & "                        UNION ALL"
    sStr = sStr & "                        SELECT ACID, TRXCD, SUBSTR(TRXNM,1,LENGTH(TRXNM)-5) AS LSNNM,"
    sStr = sStr & "                               KAEYOL, B.NUM"
    sStr = sStr & "                          FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM, CUTA, CUTB"
    sStr = sStr & "                                  FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),1,2) AS CUTA,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),4,2) AS CUTB"
    sStr = sStr & "                                          FROM SDTRX01TB"
    sStr = sStr & "                                         WHERE ACID = '" & basModule.schcd & "'"
    sStr = sStr & "                                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                                       )"
    sStr = sStr & "                                 WHERE LTRIM(CUTA,'0123456789') IS NULL"
    sStr = sStr & "                                   AND LTRIM(CUTB,'0123456789') IS NULL"
    sStr = sStr & "                                ) A,"
    sStr = sStr & "                               SDTRX90TB B"
    sStr = sStr & "                         WHERE B.NUM BETWEEN CUTA AND CUTB"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       (SELECT ACID, TRXCD, KAEYOL, LESSON, WEEKS"
    sStr = sStr & "                          FROM SDTRX11TB"
    sStr = sStr & "                         WHERE ACID  = '" & basModule.schcd & "'"
    sStr = sStr & "                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                        ) B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "                   AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "                ) A,"
    sStr = sStr & "               (SELECT LSNCD, MAX(LSNNM) AS LSNNM,"
    sStr = sStr & "                       KAEYOL, MAX(GAEYUL) AS GAEYUL,"
    sStr = sStr & "                       MAX(CLASSNM) AS CLASSNM, MAX(DAMIM) AS DAMIM,"
    sStr = sStr & "                       LSNCDNM, TCRNM, SUBJNM"
    sStr = sStr & "                  FROM (SELECT LSNCD, LSNNM, KAEYOL, GAEYUL, CLASSNM, DAMIM, LSNCDNM, TCRNM, SUBJNM"
    sStr = sStr & "                          FROM (SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                                       B.KAEYOL,"
    sStr = sStr & "                                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS GAEYUL,"
    sStr = sStr & "                                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                                       B.DAMIM,"
    sStr = sStr & "                                       B.LSNCDNM,"
    sStr = sStr & "                                       A.TCRNM, A.SUBJNM"
    sStr = sStr & "                                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                                           AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                                        ) A,"
    sStr = sStr & "                                       SDLSN01TB B"
    sStr = sStr & "                                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                                   AND A.ACID  = '" & basModule.schcd & "'"
    sStr = sStr & "                                UNION ALL"
    sStr = sStr & "                                SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                                       B.KAEYOL,"
    sStr = sStr & "                                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS GAEYUL,"
    sStr = sStr & "                                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                                       B.DAMIM,"
    sStr = sStr & "                                       B.LSNCDNM,"
    sStr = sStr & "                                       A.TCRNM, A.SUBJNM"
    sStr = sStr & "                                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                                           AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                                        ) A,"
    sStr = sStr & "                                       SDLSN02TB B"
    sStr = sStr & "                                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                                   AND A.ACID  = '" & basModule.schcd & "'"
    sStr = sStr & "                                UNION ALL"
    sStr = sStr & "                                SELECT '00000' AS LSNCD, PRT_LSNNM AS LSNNM,"
    sStr = sStr & "                                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
    sStr = sStr & "                                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','인문계','2','자연계','기타') AS GAEYUL,"
    sStr = sStr & "                                       '' AS CLASSNM,"
    sStr = sStr & "                                       '' AS DAMIM,"
    sStr = sStr & "                                       'XX' AS LSNCDNM,"
    sStr = sStr & "                                       B.TCRNM, B.SUBJNM"
    sStr = sStr & "                                  FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                                   AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                                   AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                                   AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                                   AND A.LSNCD  = '00000'"
    sStr = sStr & "                               )"
    sStr = sStr & "                       )"
    sStr = sStr & "                 GROUP BY LSNCD, KAEYOL, LSNCDNM, TCRNM, SUBJNM"
    sStr = sStr & "               ) B"
    sStr = sStr & "         WHERE A.KAEYOL  = B.KAEYOL"
    sStr = sStr & "           AND A.LSNCDNM = B.LSNCDNM"
    
    ''>> 계열
    Select Case Trim(Right(cboKaeyol, 30))
        Case "ALL"
            ' no action
        Case "01", "03"
            sStr = sStr & "   AND A.KAEYOL = '01' "
        Case "02"
            sStr = sStr & "   AND A.KAEYOL = '02' "
        Case Else
            'NO ACTION
    End Select
    
    sStr = sStr & "        UNION ALL"
'/* 정규반 ## */
    sStr = sStr & "        SELECT LSNCD, LSNNM, KAEYOL, GAEYUL, CLASSNM, DAMIM, IDX, LSNCDNM, TCRNM, SUBJNM"
    sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS GAEYUL,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.LSNCDNM,"
    sStr = sStr & "                       A.TCRNM, A.SUBJNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN01TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.schcd & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       DECODE(B.KAEYOL,'01','인문계','02','자연계','03','예체능') AS GAEYUL,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.LSNCDNM,"
    sStr = sStr & "                       A.TCRNM, A.SUBJNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN02TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.schcd & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT '00000' AS LSNCD, PRT_LSNNM AS LSNNM,"
    sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
    sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','인문계','2','자연계','기타') AS GAEYUL,"
    sStr = sStr & "                       '' AS CLASSNM,"
    sStr = sStr & "                       '' AS DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       'XX' AS LSNCDNM,"
    sStr = sStr & "                       B.TCRNM, B.SUBJNM"
    sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                   AND A.ACID   = '" & basModule.schcd & "'"
    sStr = sStr & "                   AND A.LSNCD  = '00000'"
    sStr = sStr & "               )"
    sStr = sStr & "         WHERE IDX > ' ' "
    
    ''>> 계열
    Select Case Trim(Right(cboKaeyol, 30))
        Case "ALL"
            ' no action
        Case "01", "03"
            sStr = sStr & "  AND KAEYOL = '01' "
        Case "02"
            sStr = sStr & "  AND KAEYOL = '02' "
        Case Else
            'NO ACTION
    End Select
    
    sStr = sStr & "       ) "
    sStr = sStr & " ORDER BY KAEYOL, LSNCDNM"
    

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
''>> 분원
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
                sLsnCD = "":        If IsNull(.Fields("LSNCD")) = False Then sLsnCD = Trim(.Fields("LSNCD"))
                
                
                '## 데이터 체크 << 반, 교시, 요일이 맞아야 함.
                ninDex = 0
                If sLsnCD > " " Then      '-----------------------------------------------------------------------------------------------------------------------
                    If UBound(uTimeTable) = 0 Then
                        ReDim uTimeTable(1) As tTimeTable
                        
                        ninDex = 1              ' INDEX - 1     처음 index
                        
                    Else
                        For ni = 1 To UBound(uTimeTable) Step 1
                            If StrComp(uTimeTable(ni).LSNCD, sLsnCD, vbTextCompare) = 0 Then
                               
                                ninDex = ni     ' INDEX - NI    기존 등록된 내용으로 넣음
                                
                            End If
                        Next ni
                    End If
                    
                    If ninDex = 0 Then
                        ninDex = UBound(uTimeTable) + 1
                        ReDim Preserve uTimeTable(ninDex) As tTimeTable      '<< 새로운 index 생성
                    End If
                    
                    If ninDex > 0 Then
                    '>> data 등록 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        uTimeTable(ninDex).LSNCD = sLsnCD
                        
                        uTimeTable(ninDex).GAEYUL = "":     If IsNull(.Fields("GAEYUL")) = False Then uTimeTable(ninDex).GAEYUL = Trim(.Fields("GAEYUL"))
                        uTimeTable(ninDex).LSNNM = "":      If IsNull(.Fields("LSNNM")) = False Then uTimeTable(ninDex).LSNNM = Trim(.Fields("LSNNM"))
                        uTimeTable(ninDex).CLASS_NM = "":   If IsNull(.Fields("CLASSNM")) = False Then uTimeTable(ninDex).CLASS_NM = Trim(.Fields("CLASSNM"))
                        uTimeTable(ninDex).DAMIM = "":      If IsNull(.Fields("DAMIM")) = False Then uTimeTable(ninDex).DAMIM = Trim(.Fields("DAMIM"))
                        
                        nArray = 0
                        If IsNull(.Fields("IDX")) = False Then
                            nArray = CLng(.Fields("IDX"))       '< 배열위치
                            
                            If IsNull(.Fields("SUBJNM")) = False Then uTimeTable(ninDex).DATA(nArray, 1) = Trim(.Fields("SUBJNM"))
                            If IsNull(.Fields("TCRNM")) = False Then uTimeTable(ninDex).DATA(nArray, 2) = Trim(.Fields("TCRNM"))
                        End If
                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    End If
                    
                End If      '## If sLsnCD > " " Then ---------------------------------------------------------------------------------------------------------------
                
                .MoveNext
            Next nRec       '## recordcount
        End If
    End With
            
    
    '## 모든 데이터는 전역변수 처리되어 있음.
    Call Disp_TimeTable_All_Data(1)
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    VScroll1.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "시간표 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 조회"
End Sub


'## 시간표 데이터 화면으로 view
Private Sub Disp_TimeTable_All_Data(ByVal aindex As Long)
    
    Dim UsrCtl      As Control
    Dim nRec        As Long
    
    If UBound(uTimeTable) = 0 Then
        MsgBox "시간표를 조회하세요.", vbExclamation + vbOKOnly, "시간표 조회"
        Exit Sub
    End If
    
    If UBound(uTimeTable) < aindex Or aindex < 1 Then
        MsgBox "더이상 조회할 시간표가 없습니다.", vbExclamation + vbOKOnly, "시간표 조회"
        Exit Sub
    End If
    
    VScroll1.Min = 1
    VScroll1.Max = UBound(uTimeTable)
    VScroll1.Enabled = True
    
    'ainDex의 자료만 보여줌
    If UBound(uTimeTable) >= aindex Then
    
        txtPage.Text = Trim(CStr(aindex)) & "/" & Trim(CStr(UBound(uTimeTable)))
    
        '>> 초기화
        For Each UsrCtl In Me
            With UsrCtl
                If UCase(UsrCtl.Name) = "LS" Or _
                   UCase(UsrCtl.Name) = "LT" Or _
                   UCase(UsrCtl.Name) = "RS" Or _
                   UCase(UsrCtl.Name) = "RT" Or _
                   UCase(UsrCtl.Name) = "LHD" Or _
                   UCase(UsrCtl.Name) = "RHD" Then
                    .Text = ""
                End If
            End With
        Next
    
        With uTimeTable(aindex)
        
        '// 1. header
            LHD(0).Text = "계열 : " & .GAEYUL:       RHD(0).Text = "계열 : " & .GAEYUL
            LHD(1).Text = "반 : " & .LSNNM:          RHD(1).Text = "반 : " & .LSNNM
            LHD(2).Text = "교실 : " & .CLASS_NM:     RHD(2).Text = "교실 : " & .CLASS_NM
            LHD(3).Text = "담당 : " & .DAMIM:        RHD(3).Text = "담당 : " & .DAMIM
        
        '// 2. 시간표 및 안내는 조회시 모두 처리됨.
        
        '// 3. 시간표 세부내역
            For nRec = 1 To UBound(.DATA) Step 1
                If .DATA(nRec, 1) > " " Then
                    LS(nRec).Text = .DATA(nRec, 1):      RS(nRec).Text = .DATA(nRec, 1)
                    LT(nRec).Text = .DATA(nRec, 2):      RT(nRec).Text = .DATA(nRec, 2)
                    
                End If
            Next nRec
        
        End With
    End If
    
End Sub






'>> scroll 이동
Private Sub VScroll1_Change()
    If Me.Tag = "LOAD" Then Exit Sub
    
    VScroll1.Enabled = False
        Call Disp_TimeTable_All_Data(VScroll1.value)
    VScroll1.Enabled = True
    
End Sub

Private Sub cmdShiftLeft_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS - 1) >= 1 Then
            VScroll1.value = nS - 1
            VScroll1.Enabled = False
                Call Disp_TimeTable_All_Data(VScroll1.value)
            VScroll1.Enabled = True
        End If
    End If
End Sub

Private Sub cmdShiftRight_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS + 1) <= nE Then
            VScroll1.value = nS + 1
            VScroll1.Enabled = False
                Call Disp_TimeTable_All_Data(VScroll1.value)
            VScroll1.Enabled = True
        End If
    End If
End Sub




'## 반 조회
Private Sub txtLsn_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF10
            sprLsn.Visible = False
        
            txtLsn(1).Text = ""
            Call Find_LsnData
            
        Case vbKeyCancel
            sprLsn.Visible = False
            sprTime.Visible = False
            sprinFo.Visible = False
            
        Case vbKeyBack
            txtLsn(1).Text = ""
            
    End Select
    
End Sub

Private Sub Frame1_Click()
    sprLsn.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
    
End Sub

Private Sub Frame2_Click()
    sprLsn.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False

End Sub

Private Sub txtLsn_Click(Index As Integer)
'    sprLsn.Visible = False
'    sprTime.Visible = False
'    sprinFo.Visible = False

End Sub

'반 조회
Private Sub txtLsn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
            sprLsn.Visible = False
        
            txtLsn(1).Text = ""
            Call Find_LsnData
            
    End Select
End Sub


Private Sub Find_LsnData()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Long
    Dim nRec        As Long
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    sprLsn.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "      SELECT LSNCD, LSNNM"
    sStr = sStr & "        From SDLSN01TB"
    sStr = sStr & "       WHERE ACID = '" & Trim(basModule.schcd) & "'"
    If Trim(txtLsn(0).Text) = "" Then
        sStr = sStr & "     AND LSNNM LIKE '%" & Trim(txtLsn(0).Text) & "%'"
    End If

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
''>> 분원
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
                sprLsn.MaxRows = sprLsn.MaxRows + 1
                sprLsn.Row = sprLsn.MaxRows
                
                sprLsn.Col = 1
                    sTmp = " ":     If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprLsn.Col = 2
                    sTmp = " ":     If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                
                .MoveNext
            Next nRec       '## recordcount
            
            sprLsn.Visible = True

        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "반 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "반 조회"
End Sub

'반 선택
Private Sub sprLsn_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprLsn
        .Row = Row
        .Col = 1
            txtLsn(1).Text = Trim(.Text)
        .Col = 2
            txtLsn(0).Text = Trim(.Text)
    End With
    
    sprLsn.Visible = False
End Sub





















'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 출  력
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'## 전체출력
Private Sub cmdPrintAll_Click()

    Dim nRec        As Long
    Dim bChk        As Boolean

    If UBound(uTimeTable) < 1 Then
        MsgBox "시간표 출력할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 전체반 출력"
        Exit Sub
    End If
    
    On Error GoTo ErrPrint
    
    bChk = False
    With dlgPrint
        .CancelError = True
        .ShowPrinter
        
        bChk = True
    End With
    
ErrPrint:
    If bChk = False Then
        MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "시간표 전체반 출력"
        Exit Sub
    End If
    
    On Error GoTo 0
    On Error GoTo ErrStmt
    
    nRec = 0
    cmdPrint.Tag = "ALL"
    
    Do
        nRec = nRec + 1
        txtPage.Text = "1" & "/" & Trim(CStr(UBound(uTimeTable)))
        
        Call Disp_TimeTable_All_Data(nRec)                      '<< 시간표 조회내역 보이기
        
        
        
        Me.Tag = "LOAD"
            VScroll1.value = nRec
            Call CmdPrint_Click:        DoEvents                '<< 현재 조회된 시간표 출력
            
        Me.Tag = ""

    Loop Until nRec = UBound(uTimeTable)
    
    cmdPrint.Tag = ""
    MsgBox "시간표 출력하였습니다.", vbInformation + vbOKOnly, "시간표 전체반 출력"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    cmdPrint.Tag = ""
    
    MsgBox "시간표 출력시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 전체반 출력"
    
End Sub

'## 현재 페이지만 출력
Public Sub CmdPrint_Click()

    Dim i           As Integer
    Dim X           As Integer
    Dim Y           As Integer
    Dim pRate       As Double


    Dim bChk        As Boolean


'    If UBound(uTimeTable) < 1 Then
'        MsgBox "시간표 출력할 내용이 없습니다.", vbExclamation + vbOKOnly, "시간표 출력"
'        Exit Sub
'    End If
    
    On Error GoTo 0
    On Error GoTo ErrPrint
    
    '<< 현재 페이지만 출력하면,
    If cmdPrint.Tag = "" Then
        bChk = False
        With dlgPrint
            .CancelError = True
            .ShowPrinter
            
            bChk = True
        End With
        
ErrPrint:
        If bChk = False Then
            MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "시간표 출력"
            Exit Sub
        End If
    End If
    
    On Error GoTo 0
    On Error Resume Next        '<< 에러가 나도 진행시킴
    
    '****************************************************************************************
    ' 프린터 출력초기화를 한다.
    ' PrintStartDoc (Width, Height, PaperSize, Orientation,TopMargin,LeftMargin
    '****************************************************************************************
    pRate = 1.15
    basFunction.PrintStartDoc pReportViewer.Width * pRate, pReportViewer.Height * pRate, vbPRPSA4, vbPRORLandscape, 1, 1


    '********************************************************************
    '  컬렉션을 이용하여 CONTROL을 배열로 처리한다.
    ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '  ※ 아래의 순서를 절대루 바꾸지 말것....   boss
    '********************************************************************
    Dim UsrCtl      As Control

    For Each UsrCtl In Me
        With UsrCtl

             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS") Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                 Printer.DrawWidth = 0                      ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent        ' 단색
                 Printer.FillColor = basModule.WhiteColor   ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS2") Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                 Printer.DrawWidth = 0                   ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent     ' 단색
                 Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
        End With
    Next

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS") Then
                '********************************************************************
                '  line를 이용한 box만들기(기본적으로 shape는 출력시 line를 이용한다)
                '********************************************************************
                 Printer.DrawWidth = 12
                 PrintBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
             End If
        End With
    Next


    For Each UsrCtl In Me
        With UsrCtl
             Select Case UCase(TypeName(UsrCtl))
                    Case "LINE"
                         '********************************************************************
                         '  박스/line를 긋는다.
                         '********************************************************************
                          Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                          Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                          Printer.FillStyle = vbFSTransparent
                          PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate

                    Case "LABEL"
                          '********************************************************************
                          '  Label을 그대로 출력 한다(속성)
                          '  단) transparent는 true로 처리하고 실행한다.
                          '  SetBkMode(Printer.hdc, TRANSPARENT)문장은 MS버그를 처리하기 위함
                          '********************************************************************
                          If (.Name <> "NonPrintLbl") Then
                                Printer.FontTransparent = True
                                iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
                                Printer.Font.Name = .Font.Name
                                Printer.Font.Size = .Font.Size
                                Printer.FontBold = .FontBold
                                Printer.FillColor = .BackColor
                                PrintCurrentX .Left * pRate
                                PrintCurrentY .Top * pRate
                                PrinterPrint .Caption
                                Printer.FontTransparent = False
                          End If

                    Case "TEXTBOX"
                         '********************************************************************
                         '  데이터 출력 (DATA는 TEXTBOX로 처리 한다.)
                         '********************************************************************
                          Select Case UCase(.Name)
                            Case "TXTLSN", "TXTPAGE"
                            
                            Case Else
                                Printer.Font.Name = .Font.Name
                                Printer.Font.Size = .Font.Size
                                Printer.FontBold = .FontBold
                                Printer.FillColor = .BackColor
                                PrintCurrentX .Left * pRate
                                PrintCurrentY .Top * pRate
                                PrinterPrint .Text
                         End Select
                    Case "IMAGE"
                          '********************************************************************
                          '  이미지출력 : picture 일경우 찍음
                          '********************************************************************
'                          If (object.Picture <> 0) Then
'                              Printer.FontTransparent = True
'                              iBKMode = SetBkMode(Printer.hDC, OPAQUE)
'                              ' iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
'                              PrintPicture .Picture, .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
'                          End If
             End Select
        End With
    Next

    Printer.EndDoc     ' 프린터로 보낸다

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<







