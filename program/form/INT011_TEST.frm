VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form INT011_TEST 
   Caption         =   "INT011_TEST"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form4"
   ScaleHeight     =   10440
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   124
      Top             =   0
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   125
         Top             =   30
         Width           =   14385
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13380
            TabIndex        =   135
            Text            =   "txtPage"
            Top             =   30
            Width           =   675
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "◀"
            Height          =   375
            Left            =   13020
            TabIndex        =   134
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "▶"
            Height          =   375
            Left            =   14070
            TabIndex        =   133
            Top             =   30
            Width           =   315
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "전체page 출력"
            Height          =   375
            Left            =   11580
            TabIndex        =   132
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "현재page 출력"
            Height          =   375
            Left            =   10110
            TabIndex        =   131
            Top             =   30
            Width           =   1365
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   6000
            Style           =   2  '드롭다운 목록
            TabIndex        =   130
            Top             =   67
            Width           =   915
         End
         Begin VB.TextBox txtStdNM 
            Height          =   285
            Left            =   7350
            TabIndex        =   129
            Text            =   "txtStdNM"
            Top             =   75
            Width           =   855
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "학생조회(&F)"
            Height          =   375
            Left            =   135
            TabIndex        =   128
            Top             =   45
            Width           =   1215
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   4710
            Style           =   2  '드롭다운 목록
            TabIndex        =   127
            Top             =   67
            Width           =   855
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   9210
            Style           =   2  '드롭다운 목록
            TabIndex        =   126
            Top             =   67
            Width           =   885
         End
         Begin EditLib.fpMask fpExmID_S 
            Height          =   285
            Left            =   2070
            TabIndex        =   136
            Top             =   75
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
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
            Mask            =   "AAAAAA"
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
         Begin EditLib.fpMask fpExmID_E 
            Height          =   285
            Left            =   3150
            TabIndex        =   137
            Top             =   75
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
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
            Mask            =   "AAAAAA"
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
            Left            =   5610
            TabIndex        =   143
            Top             =   120
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "출력물"
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
            Index           =   2
            Left            =   225
            TabIndex        =   142
            Top             =   180
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "수험번호        부터"
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
            Index           =   1
            Left            =   1350
            TabIndex        =   141
            Top             =   120
            Width           =   3285
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "학생"
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
            Left            =   6990
            TabIndex        =   140
            Top             =   120
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "무/유시험"
            Height          =   210
            Index           =   4
            Left            =   3690
            TabIndex        =   139
            Top             =   105
            Width           =   975
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "인터넷/학원"
            Height          =   210
            Index           =   5
            Left            =   8100
            TabIndex        =   138
            Top             =   135
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox pReportControl 
      Height          =   9855
      Left            =   0
      ScaleHeight     =   9795
      ScaleWidth      =   14415
      TabIndex        =   0
      Top             =   540
      Width           =   14475
      Begin VB.PictureBox pReportViewer 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   9825
         Left            =   0
         ScaleHeight     =   9825
         ScaleWidth      =   14175
         TabIndex        =   2
         Top             =   -30
         Width           =   14175
         Begin VB.TextBox txtMu_Type 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
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
            Height          =   405
            Left            =   7530
            TabIndex        =   158
            Top             =   9090
            Width           =   585
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   2115
            Left            =   2490
            TabIndex        =   148
            Top             =   30
            Visible         =   0   'False
            Width           =   9285
            Begin VB.TextBox Text2 
               Appearance      =   0  '평면
               BorderStyle     =   0  '없음
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2280
               TabIndex        =   152
               Text            =   "한국사, 한국지리"
               Top             =   1290
               Width           =   4245
            End
            Begin VB.TextBox 선택_수리영역 
               Appearance      =   0  '평면
               BorderStyle     =   0  '없음
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2280
               TabIndex        =   151
               Text            =   "(심화) 고난도 구문&독해"
               Top             =   900
               Width           =   8325
            End
            Begin VB.TextBox 선택_외국어 
               Appearance      =   0  '평면
               BorderStyle     =   0  '없음
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2280
               TabIndex        =   150
               Text            =   "(심화) 도형과 함수"
               Top             =   510
               Width           =   4515
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  '평면
               BorderStyle     =   0  '없음
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2310
               TabIndex        =   149
               Text            =   "(심화) 문학 독해"
               Top             =   150
               Width           =   4125
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   48
               X1              =   360
               X2              =   11010
               Y1              =   420
               Y2              =   420
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   45
               X1              =   360
               X2              =   11010
               Y1              =   810
               Y2              =   810
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   44
               X1              =   360
               X2              =   11010
               Y1              =   1170
               Y2              =   1170
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   25
               X1              =   2190
               X2              =   2190
               Y1              =   0
               Y2              =   1590
            End
            Begin VB.Label Labels 
               Alignment       =   2  '가운데 맞춤
               BackStyle       =   0  '투명
               Caption         =   "클리닉 | 탐구 선택"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   8.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1485
               Index           =   73
               Left            =   0
               TabIndex        =   157
               Top             =   90
               Width           =   195
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "국 어 (택 1)"
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
               Index           =   71
               Left            =   810
               TabIndex        =   156
               Top             =   150
               Width           =   1095
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "수 학 (택 1)"
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
               Index           =   70
               Left            =   810
               TabIndex        =   155
               Top             =   540
               Width           =   1095
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "영 어 (택 1)"
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
               Index           =   69
               Left            =   810
               TabIndex        =   154
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "사회/과학 탐구 (택 1)"
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
               Left            =   450
               TabIndex        =   153
               Top             =   1320
               Width           =   1665
            End
         End
         Begin VB.TextBox 유시험_총점 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   36
            Text            =   "100"
            Top             =   9480
            Width           =   375
         End
         Begin VB.TextBox 유시험_수학 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   35
            Text            =   "100"
            Top             =   9210
            Width           =   375
         End
         Begin VB.TextBox 유시험_영어 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   34
            Text            =   "100"
            Top             =   8910
            Width           =   375
         End
         Begin VB.TextBox 학생우편번호 
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
            Height          =   195
            Left            =   2190
            TabIndex        =   33
            Text            =   "(100-100)"
            Top             =   3510
            Width           =   1005
         End
         Begin VB.TextBox 보호자우편번호 
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
            Height          =   195
            Left            =   2190
            TabIndex        =   32
            Text            =   "(100-100)"
            Top             =   5430
            Width           =   1005
         End
         Begin VB.TextBox 선택_사회탐구 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   31
            Text            =   "현대사,세계사,경제"
            Top             =   6660
            Width           =   8325
         End
         Begin VB.TextBox 선택_과학탐구 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3000
            TabIndex        =   30
            Text            =   "물리II,생물II,지학II"
            Top             =   7440
            Width           =   4545
         End
         Begin VB.TextBox 보호자성명 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   29
            Text            =   "홍길동"
            Top             =   5055
            Width           =   1545
         End
         Begin VB.TextBox 보호자관계 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6660
            TabIndex        =   28
            Text            =   "부모"
            Top             =   5040
            Width           =   555
         End
         Begin VB.TextBox 보호자주소1 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   27
            Text            =   "서울 중구 신당동 떡복이집..................."
            Top             =   5625
            Width           =   5055
         End
         Begin VB.TextBox 보호자직업 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   26
            Text            =   "삼호물산주식회사"
            Top             =   5535
            Width           =   2955
         End
         Begin VB.TextBox 보호자연락처_휴대폰 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10200
            TabIndex        =   25
            Text            =   "999-9999-9999"
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox 보호자연락처_직장 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   24
            Text            =   "02-2104-8600"
            Top             =   6000
            Width           =   1365
         End
         Begin VB.TextBox 보호자주소2 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   23
            Text            =   "서울 중구 신당동 떡복이집..................."
            Top             =   6015
            Width           =   5055
         End
         Begin VB.TextBox 학생주소1 
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
            Height          =   195
            Left            =   2190
            TabIndex        =   22
            Text            =   "서울 송파구 삼전동"
            Top             =   3705
            Width           =   2480
         End
         Begin VB.TextBox 학생출신고 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   21
            Text            =   "학교출신교"
            Top             =   4095
            Width           =   1995
         End
         Begin VB.TextBox 졸업년도 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5400
            TabIndex        =   20
            Text            =   "2005"
            Top             =   4095
            Width           =   495
         End
         Begin VB.TextBox 학생이메일 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   19
            Text            =   "iiiboss_12345@mail.naver.com"
            Top             =   4545
            Width           =   2955
         End
         Begin VB.TextBox 학생연락처_집 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   18
            Text            =   "02-2104-8600"
            Top             =   3615
            Width           =   2955
         End
         Begin VB.TextBox 학생연락처_휴대폰 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   17
            Text            =   "011-9490-8607"
            Top             =   4095
            Width           =   2955
         End
         Begin VB.TextBox 학생주소2 
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
            Height          =   195
            Left            =   4710
            TabIndex        =   16
            Text            =   "53-21 쌍용빌라 나동 201호 "
            Top             =   3705
            Width           =   2800
         End
         Begin VB.TextBox 학생성명 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   15
            Text            =   "홍길동"
            Top             =   3135
            Width           =   1545
         End
         Begin VB.TextBox 성별 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6300
            TabIndex        =   14
            Text            =   "남자"
            Top             =   3135
            Width           =   645
         End
         Begin VB.TextBox 생년월일 
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   13
            Text            =   "9999 - 99 - 99"
            Top             =   3150
            Width           =   2955
         End
         Begin VB.TextBox 수험번호 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11910
            TabIndex        =   12
            Text            =   "N12501"
            Top             =   930
            Width           =   1005
         End
         Begin VB.TextBox 수리 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9165
            TabIndex        =   11
            Text            =   "100"
            Top             =   9345
            Width           =   375
         End
         Begin VB.TextBox 영어 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10065
            TabIndex        =   10
            Text            =   "100"
            Top             =   9345
            Width           =   375
         End
         Begin VB.TextBox 접수계열 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11505
            TabIndex        =   9
            Text            =   "예.체능계"
            Top             =   540
            Width           =   1800
         End
         Begin VB.TextBox 언어 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8355
            TabIndex        =   8
            Text            =   "100"
            Top             =   9345
            Width           =   375
         End
         Begin VB.TextBox 접수계열2 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   645
            TabIndex        =   7
            Text            =   "접수계열2"
            Top             =   2370
            Width           =   1515
         End
         Begin VB.TextBox 학원접수 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12900
            TabIndex        =   6
            Text            =   "-인"
            Top             =   930
            Width           =   495
         End
         Begin VB.TextBox 지원학원 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "돋움체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   11580
            TabIndex        =   5
            Text            =   "K"
            Top             =   930
            Width           =   315
         End
         Begin VB.TextBox 보호자연락처 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   4
            Text            =   "보호자 전화번호"
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox 제2지망 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11685
            TabIndex        =   3
            Text            =   "노량진"
            Top             =   2160
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   795
            Index           =   9
            Left            =   510
            Top             =   8880
            Width           =   15
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "(학생의)"
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
            Index           =   81
            Left            =   5850
            TabIndex        =   123
            Top             =   5070
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   24
            X1              =   7515
            X2              =   7515
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "년 2월 졸업(예정)"
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
            Index           =   80
            Left            =   6000
            TabIndex        =   122
            Top             =   4095
            Width           =   1365
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "고등학교"
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
            Index           =   79
            Left            =   4230
            TabIndex        =   121
            Top             =   4095
            Width           =   675
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   $"INT011_TEST.frx":0000
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
            Left            =   510
            TabIndex        =   120
            Top             =   8190
            Width           =   13035
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   $"INT011_TEST.frx":00A5
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
            Left            =   510
            TabIndex        =   119
            Top             =   8010
            Width           =   12975
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "※ 비  고"
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
            Left            =   12270
            TabIndex        =   118
            Top             =   5310
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "관     계"
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
            Left            =   4890
            TabIndex        =   117
            Top             =   5070
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "성     별"
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
            Left            =   4890
            TabIndex        =   116
            Top             =   3150
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "2013년 학적카드"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   24
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   64
            Left            =   510
            TabIndex        =   115
            Top             =   1620
            Width           =   4065
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "출 신 교"
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
            Left            =   1260
            TabIndex        =   114
            Top             =   4095
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "※수험번호"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   61
            Left            =   9990
            TabIndex        =   113
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "2013년"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   62
            Left            =   480
            TabIndex        =   112
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "대성학원 입학원서"
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
            Index           =   60
            Left            =   1600
            TabIndex        =   111
            Top             =   750
            Width           =   3585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   " 직 장 전 화"
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
            Index           =   59
            Left            =   7710
            TabIndex        =   110
            Top             =   6030
            Width           =   855
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   " 직업(근무처)"
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
            Index           =   58
            Left            =   7680
            TabIndex        =   109
            Top             =   5550
            Width           =   975
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "  *전화(휴대폰)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   6.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   7680
            TabIndex        =   108
            Top             =   5070
            Width           =   975
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*E-mail"
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
            Left            =   7770
            TabIndex        =   107
            Top             =   4560
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*휴 대 폰"
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
            Left            =   7770
            TabIndex        =   106
            Top             =   4110
            Width           =   700
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*전    화"
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
            Left            =   7770
            TabIndex        =   105
            Top             =   3630
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "점"
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
            Left            =   13230
            TabIndex        =   104
            Top             =   9495
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "점"
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
            Index           =   51
            Left            =   13230
            TabIndex        =   103
            Top             =   9225
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "점"
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
            Index           =   50
            Left            =   13230
            TabIndex        =   102
            Top             =   8925
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "총 점"
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
            Index           =   49
            Left            =   12090
            TabIndex        =   101
            Top             =   9495
            Width           =   405
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "영 어"
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
            Index           =   48
            Left            =   12090
            TabIndex        =   100
            Top             =   8925
            Width           =   405
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "수 학"
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
            Left            =   12090
            TabIndex        =   99
            Top             =   9225
            Width           =   405
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '투명
            Caption         =   "(인)"
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
            Left            =   10950
            TabIndex        =   98
            Top             =   9330
            Width           =   285
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '투명
            Caption         =   "(확인)"
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
            Left            =   10875
            TabIndex        =   97
            Top             =   8985
            Width           =   465
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "외국어"
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
            Left            =   10005
            TabIndex        =   96
            Top             =   8910
            Width           =   525
         End
         Begin VB.Label 수리선택 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "수리"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9060
            TabIndex        =   95
            Top             =   8970
            Width           =   705
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "11"
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
            Index           =   39
            Left            =   6720
            TabIndex        =   94
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "10"
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
            Index           =   38
            Left            =   6240
            TabIndex        =   93
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "9"
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
            Left            =   5760
            TabIndex        =   92
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "8"
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
            Left            =   5280
            TabIndex        =   91
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "7"
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
            Left            =   4800
            TabIndex        =   90
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "6"
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
            Left            =   4280
            TabIndex        =   89
            Top             =   9015
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "5"
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
            Left            =   3720
            TabIndex        =   88
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "4"
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
            Left            =   3240
            TabIndex        =   87
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "3"
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
            Index           =   31
            Left            =   2760
            TabIndex        =   86
            Top             =   9000
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "2"
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
            Index           =   30
            Left            =   2250
            TabIndex        =   85
            Top             =   9000
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "반"
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
            Index           =   28
            Left            =   720
            TabIndex        =   84
            Top             =   9405
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "월"
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
            Left            =   720
            TabIndex        =   83
            Top             =   9015
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "과학탐구 (택2)"
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
            Left            =   1710
            TabIndex        =   82
            Top             =   7470
            Width           =   1185
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "사회탐구 (택2)"
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
            Left            =   1710
            TabIndex        =   81
            Top             =   6660
            Width           =   1185
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "계"
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
            Left            =   1230
            TabIndex        =   80
            Top             =   7680
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "연"
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
            Index           =   21
            Left            =   1230
            TabIndex        =   79
            Top             =   7440
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "자"
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
            Index           =   20
            Left            =   1230
            TabIndex        =   78
            Top             =   7200
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "계"
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
            Index           =   19
            Left            =   1230
            TabIndex        =   77
            Top             =   6930
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "문"
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
            Index           =   18
            Left            =   1230
            TabIndex        =   76
            Top             =   6690
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "인"
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
            Left            =   1230
            TabIndex        =   75
            Top             =   6450
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*주    소"
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
            Left            =   1260
            TabIndex        =   74
            Top             =   5760
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*성    명"
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
            Left            =   1260
            TabIndex        =   73
            Top             =   5070
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*주    소"
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
            Left            =   1260
            TabIndex        =   72
            Top             =   3620
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*성    명"
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
            Index           =   11
            Left            =   1260
            TabIndex        =   71
            Top             =   3150
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "목"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   750
            TabIndex        =   70
            Top             =   7530
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "택"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   750
            TabIndex        =   69
            Top             =   6855
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "과"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   750
            TabIndex        =   68
            Top             =   7185
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "선"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   750
            TabIndex        =   67
            Top             =   6510
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "호"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   66
            Top             =   5475
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   750
            TabIndex        =   65
            Top             =   5910
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "보"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   750
            TabIndex        =   64
            Top             =   5040
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "생"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   750
            TabIndex        =   63
            Top             =   4290
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "학"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   750
            TabIndex        =   62
            Top             =   3330
            Width           =   195
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   52
            X1              =   2370
            X2              =   5040
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "학번:"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   2430
            TabIndex        =   61
            Top             =   2460
            Width           =   760
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   51
            X1              =   11970
            X2              =   13500
            Y1              =   9405
            Y2              =   9405
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   50
            X1              =   11970
            X2              =   13500
            Y1              =   9135
            Y2              =   9135
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   8160
            X2              =   10680
            Y1              =   9210
            Y2              =   9210
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   49
            X1              =   12570
            X2              =   12570
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   47
            X1              =   11550
            X2              =   11550
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   11970
            X2              =   11970
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   10665
            X2              =   10665
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   43
            X1              =   8970
            X2              =   8970
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   9810
            X2              =   9810
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   41
            X1              =   6105
            X2              =   6105
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   39
            X1              =   5085
            X2              =   5085
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   38
            X1              =   5595
            X2              =   5595
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   37
            X1              =   4065
            X2              =   4065
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   36
            X1              =   4575
            X2              =   4575
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   35
            X1              =   3045
            X2              =   3045
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   34
            X1              =   3555
            X2              =   3555
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   33
            X1              =   2025
            X2              =   2025
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   32
            X1              =   2535
            X2              =   2535
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   31
            X1              =   1050
            X2              =   1050
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   30
            X1              =   1515
            X2              =   1515
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   29
            X1              =   510
            X2              =   7050
            Y1              =   9285
            Y2              =   9285
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   27
            X1              =   11730
            X2              =   13500
            Y1              =   5550
            Y2              =   5550
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   26
            X1              =   11730
            X2              =   13500
            Y1              =   5190
            Y2              =   5190
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   23
            X1              =   2910
            X2              =   2910
            Y1              =   6330
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   22
            X1              =   1560
            X2              =   1560
            Y1              =   6330
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   21
            X1              =   4710
            X2              =   4710
            Y1              =   4890
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   20
            X1              =   5700
            X2              =   5700
            Y1              =   4890
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   19
            X1              =   4740
            X2              =   4740
            Y1              =   2970
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   18
            X1              =   5730
            X2              =   5730
            Y1              =   2970
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   17
            X1              =   2070
            X2              =   2070
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   8640
            X2              =   8640
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   15
            X1              =   7650
            X2              =   7650
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   1080
            X2              =   1080
            Y1              =   2970
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   13
            X1              =   11280
            X2              =   11280
            Y1              =   480
            Y2              =   1260
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   12
            X1              =   11280
            X2              =   13500
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   10
            X1              =   1080
            X2              =   11730
            Y1              =   7140
            Y2              =   7140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   8
            X1              =   2070
            X2              =   11730
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   7
            X1              =   1080
            X2              =   11730
            Y1              =   5370
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   6
            X1              =   1080
            X2              =   11730
            Y1              =   4410
            Y2              =   4410
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   1080
            X2              =   11730
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   1080
            X2              =   11730
            Y1              =   3450
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '점
            Index           =   3
            X1              =   525
            X2              =   14145
            Y1              =   8805
            Y2              =   8805
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '점
            Index           =   2
            X1              =   510
            X2              =   14130
            Y1              =   1380
            Y2              =   1380
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   0
            X1              =   510
            X2              =   11700
            Y1              =   6330
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   28
            X1              =   510
            X2              =   11730
            Y1              =   4890
            Y2              =   4890
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   825
            Index           =   5
            Left            =   510
            Top             =   8865
            Width           =   13005
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   585
            Index           =   2
            Left            =   510
            Top             =   2250
            Width           =   1755
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   4965
            Index           =   0
            Left            =   525
            Top             =   2955
            Width           =   13005
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   795
            Index           =   1
            Left            =   9960
            Top             =   480
            Width           =   3555
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "험"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   82
            Left            =   11670
            TabIndex        =   60
            Top             =   9405
            Width           =   225
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "시"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   11670
            TabIndex        =   59
            Top             =   9195
            Width           =   225
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "유"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   86
            Left            =   11670
            TabIndex        =   58
            Top             =   8985
            Width           =   225
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "언어"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   8430
            TabIndex        =   57
            Top             =   8970
            Width           =   375
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   40
            X1              =   7050
            X2              =   7050
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "12"
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
            Index           =   29
            Left            =   1200
            TabIndex        =   56
            Top             =   9000
            Width           =   195
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '투명
            Caption         =   "▶ 여유"
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
            Left            =   510
            TabIndex        =   55
            Top             =   8400
            Width           =   13065
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "1"
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
            Index           =   41
            Left            =   1720
            TabIndex        =   54
            Top             =   9000
            Width           =   195
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   53
            X1              =   6585
            X2              =   6585
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   1
            X1              =   11730
            X2              =   11730
            Y1              =   2970
            Y2              =   7920
         End
         Begin VB.Image Photo 
            Height          =   2145
            Left            =   11730
            Stretch         =   -1  'True
            Top             =   3000
            Width           =   1785
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "(영어)"
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
            Left            =   10020
            TabIndex        =   53
            Top             =   9060
            Width           =   525
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   9
            X1              =   8160
            X2              =   8160
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   0
            X1              =   11520
            X2              =   11520
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "험"
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
            Left            =   7230
            TabIndex        =   52
            Top             =   9450
            Width           =   255
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "점선 아랫부분은 기재하지 마시오"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   6.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   5760
            TabIndex        =   51
            Top             =   8625
            Width           =   2745
         End
         Begin VB.Label B_LB 
            BackStyle       =   0  '투명
            Caption         =   "선   택"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   10245
            TabIndex        =   50
            Top             =   2280
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label B_LB 
            BackStyle       =   0  '투명
            Caption         =   "2 지망"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   10245
            TabIndex        =   49
            Top             =   2070
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape L_B2 
            BorderColor     =   &H00FF0000&
            Height          =   585
            Left            =   9975
            Top             =   1965
            Visible         =   0   'False
            Width           =   3555
         End
         Begin VB.Line L_B1 
            BorderColor     =   &H00FF0000&
            Visible         =   0   'False
            X1              =   11295
            X2              =   11295
            Y1              =   1965
            Y2              =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "@ 굵은선 안에만 기재하시오."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   9990
            TabIndex        =   48
            Top             =   1425
            Width           =   2850
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "*는 필수정보이고 그 외에는 선택정보입니다."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   10005
            TabIndex        =   47
            Top             =   1650
            Width           =   3525
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000005&
            Caption         =   "'11"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   5.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1080
            TabIndex        =   46
            Top             =   8880
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "지망대학"
            Height          =   255
            Left            =   1220
            TabIndex        =   45
            Top             =   4560
            Width           =   735
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000005&
            Caption         =   "1지망"
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   44
            Top             =   4560
            Width           =   500
         End
         Begin VB.Label lab_UNI1 
            BackColor       =   &H80000005&
            Caption         =   "서울대학"
            Height          =   255
            Index           =   0
            Left            =   2700
            TabIndex        =   43
            Top             =   4560
            Width           =   800
         End
         Begin VB.Label lab_Major1 
            BackColor       =   &H80000005&
            Caption         =   "정보통신학과"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   42
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000005&
            Caption         =   "2지망"
            Height          =   255
            Left            =   4800
            TabIndex        =   41
            Top             =   4560
            Width           =   495
         End
         Begin VB.Label lab_UNI2 
            BackColor       =   &H80000005&
            Caption         =   "경북대학"
            Height          =   255
            Left            =   5400
            TabIndex        =   40
            Top             =   4560
            Width           =   795
         End
         Begin VB.Label lab_Major2 
            BackColor       =   &H80000005&
            Caption         =   "사회복지학과"
            Height          =   255
            Left            =   6360
            TabIndex        =   39
            Top             =   4560
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   4710
            X2              =   4710
            Y1              =   4410
            Y2              =   4890
         End
         Begin VB.Label lab_Mu_Type 
            BackColor       =   &H80000005&
            Caption         =   "6월 평가원"
            Height          =   255
            Left            =   12650
            TabIndex        =   38
            Top             =   2640
            Width           =   885
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "*생년월일"
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
            Left            =   7770
            TabIndex        =   37
            Top             =   3150
            Width           =   765
         End
         Begin VB.Shape L_B3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   585
            Left            =   9975
            Top             =   1980
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   795
            Index           =   1
            Left            =   9960
            Top             =   480
            Width           =   1320
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   825
            Index           =   11
            Left            =   11550
            Top             =   8865
            Width           =   420
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   825
            Index           =   10
            Left            =   7050
            Top             =   8865
            Width           =   1110
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   360
            Index           =   4
            Left            =   11730
            Top             =   5190
            Width           =   1785
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   3360
            Index           =   7
            Left            =   7680
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   480
            Index           =   6
            Left            =   4710
            Top             =   4890
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   480
            Index           =   5
            Left            =   4740
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   3375
            Index           =   0
            Left            =   1080
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   1605
            Index           =   8
            Left            =   1080
            Top             =   6330
            Width           =   1830
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9765
         Left            =   14190
         TabIndex        =   1
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   3420
      Top             =   10410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1860
      Top             =   10350
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2490
      Top             =   10350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   132
      ImageHeight     =   150
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INT011_TEST.frx":0150
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   11
      X1              =   360
      X2              =   11010
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Label Labels 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "클리닉 | 탐구 선택"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Index           =   68
      Left            =   0
      TabIndex        =   147
      Top             =   0
      Width           =   195
   End
   Begin VB.Label Labels 
      BackStyle       =   0  '투명
      Caption         =   "국 어 (택 1)"
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
      Left            =   810
      TabIndex        =   146
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Labels 
      BackStyle       =   0  '투명
      Caption         =   "영 어 (택 1)"
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
      Left            =   810
      TabIndex        =   145
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Labels 
      BackStyle       =   0  '투명
      Caption         =   "사회/과학 탐구 (택 1)"
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
      Left            =   450
      TabIndex        =   144
      Top             =   1230
      Width           =   1665
   End
End
Attribute VB_Name = "INT011_TEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : INT011
'   모 듈  목 적 : 입학원서 출력 (선행반만)
'
'   작   성   일 : 2007/12/06
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit


Private Type tSTD
    SCHNO       As String
    ACID        As String
    EXMID       As String
    STDNM       As String
    Birth_ymd       As String
    
    EXMTYPE     As String
    KAEYOL      As String
    
    SEL1        As String
    SEL2        As String
    SEL3        As String
    SEL4        As String
    SEL5        As String
    
    K_NUM       As Long
    M_NUM       As Long
    E_NUM       As Long
    TOT_NUM     As Long
    
    SEL1_SCH    As String
    SEL2_SCH    As String
    
    PASS1       As String
    PASS2       As String
    PASS3       As String
    PASS4       As String
    CL_CLOSE    As String
    CY_ACNT     As String
    TOT_AMT     As Long
    
    BASE_AMT1   As Long
    BASE_AMT2   As Long
    BASE_AMT3   As Long
    BASE_AMT4   As Long
    
    BASE_AMT5   As Long
    BASE_AMT6   As Long
    BASE_AMT7   As Long
    BASE_AMT8   As Long
    
    TAMGU_AMT1  As Long
    TAMGU_AMT2  As Long
    TAMGU_AMT3  As Long
    TAMGU_AMT4  As Long
    TAMGU_AMT5  As Long
    TAMGU_AMT6  As Long
    TAMGU_AMT7  As Long
    TAMGU_AMT8  As Long
    TAMGU_AMT9  As Long
    TAMGU_AMT10 As Long
    TAMGU_AMT11 As Long
    
    SEX         As String
    
    ZIP         As String
    ADDR1       As String
    ADDR2       As String
    TEL         As String
    CEL         As String
    EMAIL       As String
    
    HIGH_SCH    As String
    GRADE_YEAR  As String
    
    PRNT_NM     As String
    PRNT_RLTN   As String
    PRNT_ZIP    As String
    PRNT_ADDR1  As String
    PRNT_ADDR2  As String
    PRNT_TEL    As String
    PRNT_CEL    As String
    PRNT_JOB    As String
    PRNT_W_TEL  As String
    
    PHOTO_PATH  As String
    R_WAY       As String
    
    ORD_NO      As String
    UNIVCD1     As String
    MAJORCD1    As String
    UNIVCD2     As String
    MAJORCD2    As String
    MU_TYPE     As String
    IMAGE_FILE  As String
    WANT_ACID   As String
    IMAGE_DIR   As String
End Type
Private uSTD() As tSTD

Private sSavePath   As String       '<< image 경로
Private nTotRec     As Long         '<< 전체 학생수
Private checkSCH As String          '<< 지역 코드

Private Const Kangnam = "/NDOC/dshw/kangnam/register/"
Private Const MKangnam = "/NDOC/dshw/mkangnam/register/"
Private Const MSongpa = "/NDOC/dshw/msongpa/register/"
Private Const Noryangjin = "/NDOC/dshw/noryangjin/register/"
Private Const Songpa = "/NDOC/dshw/songpa/register/"
Private Const MGwanghwa = "/NDOC/dshw/kwanghwamun/register/"
Private Const Busan = "/NDOC/dshw/busan/register/"

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    lab_Mu_Type.Caption = ""
    
    
    
    Me.Tag = "LOAD"
        nTotRec = 0
        Call Clear_Form_Control
        
        sSavePath = App.Path & "\PHOTO"
        If Dir(sSavePath, vbDirectory) = "" Then
            Call MkDir(sSavePath)
        End If
        
        VScroll1.Min = 1
        VScroll1.Max = 100
        VScroll1.SmallChange = 1
        VScroll1.LargeChange = 1
        VScroll1.Enabled = False
        
        Me.Width = 14550
        Me.Height = 10755
        
        fpExmID_S.Text = ""
        fpExmID_E.Text = ""
        
        checkSCH = basModule.SchCD
        
        Select Case Trim(basModule.SchCD)
            Case "N"
                Labels(23).Caption = "사회탐구 (택2)"
                Labels(26).Caption = "과학탐구 (택2)"
                
            Case "K"
                Labels(23).Caption = "사회탐구 (택2)"
                Labels(26).Caption = "과학탐구 (택2)"
                
            Case "S"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
                
            Case "P"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
                
            Case "M"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
            
            Case "W"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
            Case "Q"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
                
            Case "J"
                Labels(23).Caption = "사회탐구 (택2)"
                Labels(26).Caption = "과학탐구 (택2)"
            
            Case "B"
                Labels(23).Caption = "사회탐구 (택3)"
                Labels(26).Caption = "과학탐구 (택3)"
            
        End Select
        
        
        
        
        '>> 무/유시험
        With cboExmType
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "유시험" & Space(30) & "1"
            .AddItem "무시험" & Space(30) & "0"
            
            .ListIndex = 0
        End With
        
        '>> 계열
        With cboKaeyol
            .Clear
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            .ListIndex = 0
        End With
        
        txtStdNM.Text = ""
        
        '>> 인터넷/학원 구분
        With cboinGbn
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "인터넷" & Space(30) & "INT"
            .AddItem "학원" & Space(30) & "HAK"
            
            .ListIndex = 0
        End With
        
        '>> 선행반/ 종합반 구분
'        With cboSel
'            .Clear
'            .AddItem "선행" & Space(30) & "01"
'            '.AddItem "종합" & Space(30) & "02"
'
'            .ListIndex = 0
'        End With
'
'        ReDim uSTD(0) As tSTD
        
    Me.Tag = ""
    
    txtMu_Type.Font.Size = 6
    txtMu_Type.Text = "2013수능"
    
End Sub

Private Sub Clear_Form_Control()
    Dim UsrCtl      As Control
    For Each UsrCtl In Me
        With UsrCtl
             If UCase(TypeName(UsrCtl)) = "TEXTBOX" Then .Text = ""
             If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
             If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
        End With
    Next
    
    학생성명.Tag = ""
    수험번호.Tag = ""
    lab_UNI1(0).Caption = ""
    lab_Major1(0).Caption = ""
    lab_UNI2.Caption = ""
    lab_Major2.Caption = ""
    'Height = 3990
    'Width = 4890   ' 높이와 너비를 설정합니다.
    Set Photo.Picture = imgList.ListImages.Item(1).Picture
    
    
'>> 학년별 내역
    Labels(23).Caption = "사회탐구 (택2)"
    Labels(26).Caption = "과학탐구 (택2)"
    
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            OPTIONS(0) = "● 선행학습반 인문계는 선택과목 4과목 중 2과목을 선택할 수 있으며, 국사, 경제, 정치, 세계지리, 세계사, 법과사회, 경제지리, 제2외국어 과목은 재수 정규반부터 수업합니다."
            OPTIONS(1) = "● 선행학습반 자연계는 선택과목 4과목 중 2과목을 선택할 수 있으며, 과학 II (4과목)는 재수 정규반부터 수업합니다."
            OPTIONS(2) = "● 반당 수강생 수 증감에 따라 분반 또는 합반할 수 있습니다."
           
        Case "K", "W", "Q"
            OPTIONS(0) = ""
            OPTIONS(1) = ""
            
        Case "S"
            OPTIONS(0) = "▶ 인문계 : 사회탐구영역 선택과목 중 법과사회, 세계지리, 세계사, 경제지리 및 제2외국어는 정규반부터 수업합니다."
            OPTIONS(1) = "▶ 자연계 : 과학탐구영역 선택과목 중 과학II (4과목) 및 수리영역 선택과목 미적분은 정규반부터 수업합니다."
            OPTIONS(2) = "※선행학습반 인문계는 7과목 중 3과목을, 자연계는 4과목 중 3과목을 선택할 수 있습니다."
                  
        Case "P"
            OPTIONS(0) = ""
            OPTIONS(1) = ""
                             
        Case "M"        '> 강남 마이맥
            Labels(23).Caption = "사회탐구 (택4)"
            Labels(26).Caption = "과학탐구 (택4)"
            
            OPTIONS(0) = "▶ 인문계 사회선택은 4과목을 선택하십시요. 세계지리, 경제지리는 극소수일 경우 성반되지 않을 수 있습니다."
            OPTIONS(1) = "▶ 자연계 수리영역 선택과목 중 확률통계는 극소수일 경우 성반되지 않을 수 있습니다. 이산수학은 수업하지 않습니다."
            
        Case "J"        '> 양재
            OPTIONS(0) = "▶ 인문계 : 사회탐구영역 선택과목 중 법과사회, 세계지리, 세계사, 경제기리, 및 제2외국어는 정규반부터 수업합니다."
            OPTIONS(1) = "▶ 자연계 : 과학탐구영역 선택과목 중 과학II(4과목)는 정규반부터 수업합니다."
            OPTIONS(2) = "※선행학습반 인문계는 7과목 중 2과목을, 자연계는 4과목 중 2과목을 선택할 수 있습니다."
            
        Case "B"        '> 부산
            OPTIONS(0) = "▶ 선행학습반 인문계는 6과목 중 3과목을 선택할 수 있으며, 경제, 세계지리, 세계사, 법과사회, 경제지리, 제2외국어 과목은 정규 종합반부터 수업합니다."
            OPTIONS(1) = "▶ 선행학습반 자연계는 4과목 중 3과목을 선택할 수 있으며, 과학II(4과목), 수리영역 선택과목 적분, 확률통계는 정규 종합반부터 수업합니다."
            OPTIONS(2) = ""
            
    End Select
    
End Sub


Private Sub cmdShiftLeft_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS - 1) >= 1 Then
            VScroll1.value = nS - 1
            VScroll1.Enabled = False
                Call Std_Data_Show(VScroll1.value)
            VScroll1.Enabled = True
        End If
    End If
End Sub

Private Sub cmdShiftRight_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS + 1) <= nE Then
            VScroll1.value = nS + 1
            VScroll1.Enabled = False
                Call Std_Data_Show(VScroll1.value)
            VScroll1.Enabled = True
        End If
    End If
End Sub



'>> 학생 조회
Private Sub cmdFind_Click()
    
    On Error GoTo ErrStmt
    Me.MousePointer = vbHourglass
    
    ReDim uSTD(0) As tSTD
    
    cmdFind.Enabled = False
        Call Get_STD_Data
        
    cmdFind.Enabled = True
    
    Me.MousePointer = vbDefault
    Exit Sub
ErrStmt:
    Me.MousePointer = vbDefault
    MsgBox "학생조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
    On Error GoTo 0

End Sub

Private Sub Get_STD_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim UsrCtl      As Control
    
    
    '<< 초기 작업 : 제약조건
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ROWNUM AS ID, "
    sStr = sStr & "         SCHNO      , ACID       , EXMID      , STDNM      , SUBSTR(Birth_ymd,1,4)||'-'||SUBSTR(Birth_ymd,5,2) ||'-'||SUBSTR(Birth_ymd,7,2) AS Birth_ymd,"
    sStr = sStr & "         EXMTYPE    , KAEYOL     ,"
    sStr = sStr & "         SEL1       , SEL2       , SEL3       , SEL4       , SEL5       ,"
    sStr = sStr & "         K_NUM      , M_NUM      , E_NUM      , TOT_NUM    ,"
    sStr = sStr & "         SEL1_SCH   , SEL2_SCH   ,"
    sStr = sStr & "         PASS1      , PASS2      , PASS3      , PASS4      , CL_CLOSE   ,"
    sStr = sStr & "         CY_ACNT    , TOT_AMT    ,"
    sStr = sStr & "         BASE_AMT1  , BASE_AMT2  , BASE_AMT3  , BASE_AMT4  , "
    sStr = sStr & "         BASE_AMT5  , BASE_AMT6  , BASE_AMT7  , BASE_AMT8  , "
    sStr = sStr & "         TAMGU_AMT1 , TAMGU_AMT2 , TAMGU_AMT3 , TAMGU_AMT4 , TAMGU_AMT5 ,"
    sStr = sStr & "         TAMGU_AMT6 , TAMGU_AMT7 , TAMGU_AMT8 , TAMGU_AMT9 , TAMGU_AMT10, TAMGU_AMT11,"
    sStr = sStr & "         DECODE(SEX,'M','남','F','여') AS SEX        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS ZIP, ADDR1      , ADDR2      ,"
    sStr = sStr & "         TEL        , CEL        , EMAIL      ,"
    sStr = sStr & "         HIGH_SCH   , GRADE_YEAR ,"
    sStr = sStr & "         PRNT_NM    , DECODE(PRNT_RLTN,'1','부','2','모','3',' ') AS PRNT_RLTN, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS PRNT_ZIP, PRNT_ADDR1 , PRNT_ADDR2 ,"
    sStr = sStr & "         PRNT_TEL   , PRNT_CEL   , PRNT_JOB   , PRNT_W_TEL ,"
    sStr = sStr & "         PHOTO_PATH , DECODE(R_WAY,'1','','2','-int','3','') AS R_WAY, ORD_NO, "
    sStr = sStr & "         UNIVCD1 , MAJORCD1,  UNIVCD2 , MAJORCD2, MU_TYPE,"
    sStr = sStr & "         ACID||EXMID AS IMAGE_FILE, "
    sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID "       '< TEST
    
    '****************************** < IMAGE 저장 디렉토리 > **********************************************
    Select Case basModule.SchCD
        Case "N"        '< 노량진
            sStr = sStr & "'" & Noryangjin & "'||"
                                    
        Case "K", "W", "Q"      '< 강남
            sStr = sStr & "'" & Kangnam & "'||"
        Case "S"        '< 송파
            sStr = sStr & "'" & Songpa & "'||"
        Case "P"        '< 송파마이맥
            sStr = sStr & "'" & MSongpa & "'||"
        Case "M"        '< 강남마이맥
            sStr = sStr & "'" & MKangnam & "'||"
        Case "J"        '< 양재
            sStr = sStr & "'" & MGwanghwa & "'||"
        Case "B"        '< 부산
            sStr = sStr & "'" & Busan & "'||"
    End Select
                            sStr = sStr & "DECODE("
                                    sStr = sStr & "     KAEYOL||EXMTYPE,"
                                    sStr = sStr & "         '010','1A',"
                                    sStr = sStr & "         '011','1B',"
                                    sStr = sStr & "         '020','2A',"
                                    sStr = sStr & "         '021','2B',"
                                    sStr = sStr & "         '030','3A',"
                                    sStr = sStr & "         '031','3B',"
                                    sStr = sStr & "         '040','4A',"
                                    sStr = sStr & "         '041','4B',"
                                    sStr = sStr & "         '050','5A',"
                                    sStr = sStr & "         '051','5B',"
                                    sStr = sStr & "         '060','6A',"
                                    sStr = sStr & "         '061','6B',"
                                    sStr = sStr & "         '070','7A',"
                                    sStr = sStr & "         '071','7B',"
                                    sStr = sStr & "         '080','8A',"
                                    sStr = sStr & "         '081','8B',"
                                    sStr = sStr & "         '090','9A',"
                                    sStr = sStr & "         '091','9B')||'/'||ORD_NO||'.jpg' AS IMAGE_DIR"
    '******************************************************************************************************
    
    sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO ,  "
            sStr = sStr & "         MAX(D_UNIVCD) AS UNIVCD1 , MAX(D_MAJORCD) AS MAJORCD1,"
            sStr = sStr & "         MAX(D_UNIVCD2) AS UNIVCD2 , MAX(D_MAJORCD2) AS MAJORCD2,"
            sStr = sStr & "         MAX(MU_TYPE) AS MU_TYPE"
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- 전체학생 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            '>> 수험번호
            If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
                sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
                sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
            ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
                sStr = sStr & "         AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
                ' no action
            End If
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "         AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "         AND EXMTYPE = '1'"
            End If
            '>> 계열
            Select Case Trim(Right(cboKaeyol, 30))
                Case "XX"
                    ' no action
                Case "01", "03"
                    sStr = sStr & "     AND SEL1 > ' ' "
                Case "02"
                    sStr = sStr & "     AND SEL3 > ' ' "
            End Select
            '>> 학생명
            If Trim(txtStdNM.Text) > " " Then
                sStr = sStr & "         AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
            End If
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "         AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "         AND R_WAY IN ('1','3') "
            End If
            sStr = sStr & "             AND EXMID > ' ' "
            sStr = sStr & "             AND BIGO2 IS NULL"          '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- 전체학생 조회 END
            '---------------------------------------------------------------------------- 합격자 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "         AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "         AND EXMTYPE = '1'"
            End If
            '>> 계열
            Select Case Trim(Right(cboKaeyol, 30))
                Case "XX"
                    ' no action
                Case "01", "03"
                    sStr = sStr & "     AND SEL1 > ' ' "
                Case "02"
                    sStr = sStr & "     AND SEL3 > ' ' "
            End Select
            '>> 학생명
            If Trim(txtStdNM.Text) > " " Then
                sStr = sStr & "         AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
            End If
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "         AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "         AND R_WAY IN ('1','3') "
            End If
            
            sStr = sStr & "             AND BIGO2 IS NULL"          '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    
            sStr = sStr & "          )"
            sStr = sStr & "   GROUP BY SCHNO"
            '---------------------------------------------------------------------------- 합격자 조회 END
    
    sStr = sStr & "    ) "
    sStr = sStr & " WHERE SCHNO > ' ' "
    '>> 수험번호
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
    sStr = sStr & " ORDER BY EXMID "
    'Text1.Text = sStr
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30


'>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
''>> 수험번호
'        If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
'            sTmp = Trim(fpExmID_S.UnFmtText)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'            sTmp = Trim(fpExmID_E.UnFmtText)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
'            sTmp = Trim(fpExmID_S.UnFmtText)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
'            sTmp = Trim(fpExmID_S.UnFmtText)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
'            ' no action
'        End If
'>> 학생명
'        If Trim(txtStdNM.Text) > " " Then
'            sTmp = "%" & Trim(txtStdNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            
            ReDim uSTD(0) As tSTD
            
            For Each UsrCtl In Me
                With UsrCtl
                     If UCase(TypeName(UsrCtl)) = "TEXTBOX" Then .Text = ""
                     If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
                     If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
                End With
            Next
            
            Set Photo.Picture = imgList.ListImages.Item(1).Picture
            
            MsgBox "해당조회대상자가 없습니다.", vbExclamation + vbOKOnly, "원서출력 조회"
            
        ElseIf .RecordCount > 0 Then
            nTotRec = .RecordCount
            
            .MoveFirst
            
            ReDim uSTD(.RecordCount) As tSTD
            
            VScroll1.Max = .RecordCount
            VScroll1.Enabled = True
            
            For nRec = 1 To .RecordCount Step 1
            
                sTmp = "":      If IsNull(.Fields("SCHNO")) = False Then sTmp = .Fields("SCHNO")
                    uSTD(nRec).SCHNO = sTmp
                sTmp = "":      If IsNull(.Fields("ACID")) = False Then sTmp = .Fields("ACID")
                    uSTD(nRec).ACID = sTmp
                sTmp = "":      If IsNull(.Fields("EXMID")) = False Then sTmp = .Fields("EXMID")
                    uSTD(nRec).EXMID = sTmp
                sTmp = "":      If IsNull(.Fields("STDNM")) = False Then sTmp = .Fields("STDNM")
                    uSTD(nRec).STDNM = sTmp
                sTmp = "":      If IsNull(.Fields("Birth_ymd")) = False Then sTmp = .Fields("Birth_ymd")
                    uSTD(nRec).Birth_ymd = sTmp
                
                sTmp = "":      If IsNull(.Fields("EXMTYPE")) = False Then sTmp = .Fields("EXMTYPE")
                    uSTD(nRec).EXMTYPE = sTmp
                sTmp = "":      If IsNull(.Fields("KAEYOL")) = False Then sTmp = .Fields("KAEYOL")
                    uSTD(nRec).KAEYOL = sTmp
                
                sTmp = "":      If IsNull(.Fields("SEL1")) = False Then sTmp = .Fields("SEL1")
                    uSTD(nRec).SEL1 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL2")) = False Then sTmp = .Fields("SEL2")
                    uSTD(nRec).SEL2 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL3")) = False Then sTmp = .Fields("SEL3")
                    uSTD(nRec).SEL3 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL4")) = False Then sTmp = .Fields("SEL4")
                    uSTD(nRec).SEL4 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL5")) = False Then sTmp = .Fields("SEL5")
                    uSTD(nRec).SEL5 = sTmp
                
                nTmp = 0:      If IsNumeric(.Fields("K_NUM")) = True Then nTmp = .Fields("K_NUM")
                    uSTD(nRec).K_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("M_NUM")) = True Then nTmp = .Fields("M_NUM")
                    uSTD(nRec).M_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("E_NUM")) = True Then nTmp = .Fields("E_NUM")
                    uSTD(nRec).E_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = .Fields("TOT_NUM")
                    uSTD(nRec).TOT_NUM = nTmp
                
                
                sTmp = "":      If IsNull(.Fields("SEL1_SCH")) = False Then sTmp = .Fields("SEL1_SCH")
                    uSTD(nRec).SEL1_SCH = .Fields("SEL1_SCH") = sTmp
                    
                    Select Case uSTD(nRec).SEL1_SCH
                        Case "N"
                            uSTD(nRec).SEL1_SCH = "노량진"
                        Case "K"
                            uSTD(nRec).SEL1_SCH = "강남"
                        Case "S"
                            uSTD(nRec).SEL1_SCH = "송파"
                        Case "P"
                            uSTD(nRec).SEL1_SCH = "송파 M"
                        Case "M"
                            uSTD(nRec).SEL1_SCH = "강남 M"
                            
                        Case "W"
                            uSTD(nRec).SEL1_SCH = "주말법의대"
                        Case "Q"
                            uSTD(nRec).SEL1_SCH = "야간법의대"
                        
                        Case "J"
                            uSTD(nRec).SEL1_SCH = "양재"
                        Case "B"
                            uSTD(nRec).SEL1_SCH = "부산"
                            
                    End Select
                
                
                sTmp = "":      If IsNull(.Fields("SEL2_SCH")) = False Then sTmp = .Fields("SEL2_SCH")
                    uSTD(nRec).SEL2_SCH = sTmp
                    
                    Select Case uSTD(nRec).SEL2_SCH
                        Case "N"
                            uSTD(nRec).SEL2_SCH = "노량진"
                        Case "K"
                            uSTD(nRec).SEL2_SCH = "강남"
                        Case "S"
                            uSTD(nRec).SEL2_SCH = "송파"
                        Case "P"
                            uSTD(nRec).SEL2_SCH = "송파 M"
                        Case "M"
                            uSTD(nRec).SEL2_SCH = "강남 M"
                            
                        Case "W"
                            uSTD(nRec).SEL2_SCH = "주말법의대"
                        Case "Q"
                            uSTD(nRec).SEL2_SCH = "야간법의대"
                            
                        Case "J"
                            uSTD(nRec).SEL2_SCH = "양재"
                        Case "B"
                            uSTD(nRec).SEL2_SCH = "부산"
                        
                    End Select
                
                
                sTmp = "":      If IsNull(.Fields("PASS1")) = False Then sTmp = .Fields("PASS1")
                    uSTD(nRec).PASS1 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS2")) = False Then sTmp = .Fields("PASS2")
                    uSTD(nRec).PASS2 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS3")) = False Then sTmp = .Fields("PASS3")
                    uSTD(nRec).PASS3 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS4")) = False Then sTmp = .Fields("PASS4")
                    uSTD(nRec).PASS4 = sTmp
                    
                
                sTmp = "":      If IsNull(.Fields("CL_CLOSE")) = False Then sTmp = .Fields("CL_CLOSE")
                    uSTD(nRec).CL_CLOSE = sTmp
                sTmp = "":      If IsNull(.Fields("CY_ACNT")) = False Then sTmp = .Fields("CY_ACNT")
                    uSTD(nRec).CY_ACNT = sTmp
                nTmp = 0:       If IsNull(.Fields("TOT_AMT")) = False Then nTmp = .Fields("TOT_AMT")
                    uSTD(nRec).TOT_AMT = nTmp
                
                
                nTmp = 0:       If IsNull(.Fields("BASE_AMT1")) = False Then nTmp = .Fields("BASE_AMT1")
                    uSTD(nRec).BASE_AMT1 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT2")) = False Then nTmp = .Fields("BASE_AMT2")
                    uSTD(nRec).BASE_AMT2 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT3")) = False Then nTmp = .Fields("BASE_AMT3")
                    uSTD(nRec).BASE_AMT3 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT4")) = False Then nTmp = .Fields("BASE_AMT4")
                    uSTD(nRec).BASE_AMT4 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT5")) = False Then nTmp = .Fields("BASE_AMT5")
                    uSTD(nRec).BASE_AMT5 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT6")) = False Then nTmp = .Fields("BASE_AMT6")
                    uSTD(nRec).BASE_AMT6 = nTmp
                    
                    
                nTmp = 0:       If IsNull(.Fields("BASE_AMT7")) = False Then nTmp = .Fields("BASE_AMT7")
                    uSTD(nRec).BASE_AMT7 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT8")) = False Then nTmp = .Fields("BASE_AMT8")
                    uSTD(nRec).BASE_AMT8 = nTmp
                                                                                                                                                  
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT1")) = False Then nTmp = .Fields("TAMGU_AMT1")
                    uSTD(nRec).TAMGU_AMT1 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT2")) = False Then nTmp = .Fields("TAMGU_AMT2")
                    uSTD(nRec).TAMGU_AMT2 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT3")) = False Then nTmp = .Fields("TAMGU_AMT3")
                    uSTD(nRec).TAMGU_AMT3 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT4")) = False Then nTmp = .Fields("TAMGU_AMT4")
                    uSTD(nRec).TAMGU_AMT4 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT5")) = False Then nTmp = .Fields("TAMGU_AMT5")
                    uSTD(nRec).TAMGU_AMT5 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT6")) = False Then nTmp = .Fields("TAMGU_AMT6")
                    uSTD(nRec).TAMGU_AMT6 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT7")) = False Then nTmp = .Fields("TAMGU_AMT7")
                    uSTD(nRec).TAMGU_AMT7 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT8")) = False Then nTmp = .Fields("TAMGU_AMT8")
                    uSTD(nRec).TAMGU_AMT8 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT9")) = False Then nTmp = .Fields("TAMGU_AMT9")
                    uSTD(nRec).TAMGU_AMT9 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT10")) = False Then nTmp = .Fields("TAMGU_AMT10")
                    uSTD(nRec).TAMGU_AMT10 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT11")) = False Then nTmp = .Fields("TAMGU_AMT11")
                    uSTD(nRec).TAMGU_AMT11 = nTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("SEX")) = False Then sTmp = .Fields("SEX")
                    uSTD(nRec).SEX = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("ZIP")) = False Then sTmp = .Fields("ZIP")
                    uSTD(nRec).ZIP = sTmp
                sTmp = "":      If IsNull(.Fields("ADDR1")) = False Then sTmp = .Fields("ADDR1")
                    uSTD(nRec).ADDR1 = sTmp
                sTmp = "":      If IsNull(.Fields("ADDR2")) = False Then sTmp = .Fields("ADDR2")
                    uSTD(nRec).ADDR2 = sTmp
                sTmp = "":      If IsNull(.Fields("TEL")) = False Then sTmp = .Fields("TEL")
                    uSTD(nRec).TEL = sTmp
                sTmp = "":      If IsNull(.Fields("CEL")) = False Then sTmp = .Fields("CEL")
                    uSTD(nRec).CEL = sTmp
                sTmp = "":      If IsNull(.Fields("EMAIL")) = False Then sTmp = .Fields("EMAIL")
                    uSTD(nRec).EMAIL = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("HIGH_SCH")) = False Then sTmp = .Fields("HIGH_SCH")
                    uSTD(nRec).HIGH_SCH = sTmp
                sTmp = "":      If IsNull(.Fields("GRADE_YEAR")) = False Then sTmp = .Fields("GRADE_YEAR")
                    uSTD(nRec).GRADE_YEAR = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PRNT_NM")) = False Then sTmp = .Fields("PRNT_NM")
                    uSTD(nRec).PRNT_NM = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_RLTN")) = False Then sTmp = .Fields("PRNT_RLTN")
                    uSTD(nRec).PRNT_RLTN = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PRNT_ZIP")) = False Then sTmp = .Fields("PRNT_ZIP")
                    uSTD(nRec).PRNT_ZIP = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_ADDR1")) = False Then sTmp = .Fields("PRNT_ADDR1")
                    uSTD(nRec).PRNT_ADDR1 = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_ADDR2")) = False Then sTmp = .Fields("PRNT_ADDR2")
                    uSTD(nRec).PRNT_ADDR2 = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_TEL")) = False Then sTmp = .Fields("PRNT_TEL")
                    uSTD(nRec).PRNT_TEL = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_CEL")) = False Then sTmp = .Fields("PRNT_CEL")
                    uSTD(nRec).PRNT_CEL = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_JOB")) = False Then sTmp = .Fields("PRNT_JOB")
                    uSTD(nRec).PRNT_JOB = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_W_TEL")) = False Then sTmp = .Fields("PRNT_W_TEL")
                    uSTD(nRec).PRNT_W_TEL = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PHOTO_PATH")) = False Then sTmp = .Fields("PHOTO_PATH")
                    uSTD(nRec).PHOTO_PATH = sTmp

                sTmp = "":      If IsNull(.Fields("R_WAY")) = False Then sTmp = .Fields("R_WAY")
                    uSTD(nRec).R_WAY = sTmp
                
                sTmp = "":      If IsNull(.Fields("ORD_NO")) = False Then sTmp = .Fields("ORD_NO")
                    uSTD(nRec).ORD_NO = sTmp
                
                sTmp = "":      If IsNull(.Fields("UNIVCD1")) = False Then sTmp = .Fields("UNIVCD1")
                uSTD(nRec).UNIVCD1 = sTmp
                sTmp = "":      If IsNull(.Fields("UNIVCD2")) = False Then sTmp = .Fields("UNIVCD2")
                uSTD(nRec).UNIVCD2 = sTmp
                sTmp = "":      If IsNull(.Fields("MAJORCD1")) = False Then sTmp = .Fields("MAJORCD1")
                uSTD(nRec).MAJORCD1 = sTmp
                sTmp = "":      If IsNull(.Fields("MAJORCD2")) = False Then sTmp = .Fields("MAJORCD2")
                uSTD(nRec).MAJORCD2 = sTmp
                sTmp = "":      If IsNull(.Fields("MU_TYPE")) = False Then sTmp = .Fields("MU_TYPE")
                uSTD(nRec).MU_TYPE = sTmp
                    
                sTmp = "":      If IsNull(.Fields("IMAGE_FILE")) = False Then sTmp = .Fields("IMAGE_FILE")
                    uSTD(nRec).IMAGE_FILE = sTmp
                    
                sTmp = "":      If IsNull(.Fields("WANT_ACID")) = False Then sTmp = .Fields("WANT_ACID")
                    uSTD(nRec).WANT_ACID = sTmp
                    
                sTmp = "":      If IsNull(.Fields("IMAGE_DIR")) = False Then sTmp = .Fields("IMAGE_DIR")
                    uSTD(nRec).IMAGE_DIR = sTmp
                    
                .MoveNext
                
            Next nRec
            
            Call Get_STD_image              '<< 이미지 자료 가져오기
            
            Call Std_Data_Show(1)           '<< 학생자료 화면 보이기
            Me.Tag = "LOAD"
                VScroll1.value = 1
                
                txtPage.Text = "1/" & Trim(CStr(nTotRec))
            Me.Tag = ""
            
        End If
    End With

    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "학생조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    VScroll1.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "학생조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
End Sub




'>> scroll 이동
Private Sub VScroll1_Change()
    If Me.Tag = "LOAD" Then Exit Sub
    
    VScroll1.Enabled = False
        Call Std_Data_Show(VScroll1.value)
        txtPage.Text = Trim(CStr(VScroll1.value)) & "/" & Trim(CStr(nTotRec))
    VScroll1.Enabled = True
End Sub

Private Sub lvRecvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Me.Tag = "LOAD" Then Exit Sub
    
    Call Std_Data_Show(Item.Index)
    
End Sub

Private Sub Std_Data_Show(Index As Long)
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If UBound(uSTD) < 1 Then Exit Sub
    If UBound(uSTD) < Index Then Exit Sub
    
    With uSTD(Index)
        
        Select Case Trim(.KAEYOL)   '<< 계열: 01,02,03-인문,자연,예체   04,05-수능인문,자연  06,07 -강남법대,의대
            Case "01":  접수계열.Text = "인 문 계"
                        접수계열2.Text = "인    문"
            Case "02":  접수계열.Text = "자 연 계"
                        접수계열2.Text = "자    연"
            Case "03":  Select Case Trim(basModule.SchCD)
                               Case "S":  접수계열.Text = "신설-인문"
                                          접수계열2.Text = "인    문"
                               Case "N":  접수계열.Text = "예.체능계"
                                          접수계열2.Text = "예.체능계"
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "04":  Select Case Trim(basModule.SchCD)
                               Case "K":  접수계열.Text = "법대(주말)"
                                          접수계열2.Text = "인    문"
                               Case "S":  접수계열.Text = "신설-자연"
                                          접수계열2.Text = "자    연"
                               Case "N":  접수계열.Text = "수리(나) 자연"
                                          접수계열2.Text = "수 리 나 "
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "05":  Select Case Trim(basModule.SchCD)
                               Case "K":  접수계열.Text = "의대(주말)"
                                          접수계열2.Text = "자    연"
                               Case "S":  접수계열.Text = "수능 자연"
                                          접수계열2.Text = "수능대비"
                               Case "N":  접수계열.Text = "수능 자연"
                                          접수계열2.Text = "수능대비 "
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "06":  Select Case Trim(basModule.SchCD)
                               Case "K":  접수계열.Text = "법대(야간)"
                                          접수계열2.Text = "인    문"
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "07":  Select Case Trim(basModule.SchCD)
                               Case "K":  접수계열.Text = "의대(야간)"
                                          접수계열2.Text = "자    연"
                               Case "N":  접수계열.Text = "신설-인문"
                                          접수계열2.Text = "인    문"
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "08":  Select Case Trim(basModule.SchCD)
                               Case "N":  접수계열.Text = "신설-자연"
                                          접수계열2.Text = "자    연"
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case "09":  Select Case Trim(basModule.SchCD)
                               Case "N": 접수계열.Text = "교대-자연"
                                          접수계열2.Text = "자    연"
                               Case Else: 접수계열.Text = ""
                                          접수계열2.Text = ""
                        End Select
            Case Else:  접수계열.Text = ""
        End Select
        
        제2지망.Text = .SEL2_SCH
        
        수험번호.Text = .EXMID
        학생성명.Text = .STDNM
        성별.Text = .SEX
        생년월일.Text = .Birth_ymd
        학생우편번호.Text = "(" & .ZIP & ")"
        학생주소1.Text = .ADDR1
        학생주소2.Text = .ADDR2
        학생출신고.Text = .HIGH_SCH
        졸업년도.Text = .GRADE_YEAR
        학생이메일.Text = .EMAIL
        학생연락처_집.Text = .TEL
        학생연락처_휴대폰.Text = .CEL
        
        lab_UNI1(0).Caption = .UNIVCD1
        lab_Major1(0).Caption = .MAJORCD1
        lab_UNI2.Caption = .UNIVCD2
        lab_Major2.Caption = .MAJORCD2
        
        Select Case Trim(.MU_TYPE)
            Case "1":
            lab_Mu_Type.Caption = "수능등급"
            Case "2":
            lab_Mu_Type.Caption = "6월 평가원"
            Case "3":
            lab_Mu_Type.Caption = "9월 평가원"
        End Select
        
        보호자성명.Text = .PRNT_NM
        보호자관계.Text = .PRNT_RLTN
        
        보호자연락처.Text = .PRNT_TEL
        보호자연락처_휴대폰.Text = .PRNT_CEL
        보호자우편번호.Text = "(" & .PRNT_ZIP & ")"
        보호자주소1.Text = .PRNT_ADDR1
        보호자주소2.Text = .PRNT_ADDR2
        
        보호자직업.Text = .PRNT_JOB
        보호자연락처_직장.Text = .PRNT_W_TEL
                             
        선택_사회탐구.Text = " "
        선택_과학탐구.Text = " "
        
        선택_사회탐구.Text = Div_Gwamok_NM("SEL1", .SEL1)
        선택_과학탐구.Text = Div_Gwamok_NM("SEL3", .SEL3)
        
        언어.Text = ""
        수리.Text = ""
        영어.Text = ""
        
        유시험_수학.Text = ""
        유시험_영어.Text = ""
        유시험_총점.Text = ""
        
        Select Case Trim(.EXMTYPE)
            Case "0"
                언어.Text = .K_NUM
                수리.Text = .M_NUM
                영어.Text = .E_NUM

            Case "1"
                If .M_NUM = 0 Then
                    유시험_수학.Text = ""
                Else
                    유시험_수학.Text = .M_NUM
                End If
                
                If .E_NUM = 0 Then
                    유시험_영어.Text = ""
                Else
                    유시험_영어.Text = .E_NUM
                End If
                
                If .TOT_NUM = 0 Then
                    유시험_총점.Text = ""
                Else
                    유시험_총점.Text = .TOT_NUM
                End If
                
                '유시험_수학.Text = .M_NUM
                '유시험_영어.Text = .E_NUM
                '유시험_총점.Text = .TOT_NUM
        End Select
        
        '>> 인문계 나형, 자연계 가형
        Select Case Trim(.KAEYOL)
            Case "01", "02", "04", "05", "06", "07"
                수리선택.Caption = "수리"
            Case "03"
                수리선택.Caption = "수리"                                               '<< 예체능
            Case Else
                수리선택.Caption = ""
        End Select
        
        학생성명.Tag = .SCHNO
        수험번호.Tag = .ORD_NO
        학원접수.Text = .R_WAY
        지원학원.Text = .WANT_ACID
        
        Set Photo.Picture = CheckJPG(sSavePath & "\" & .IMAGE_FILE & ".jpg")
        
        If Trim(basModule.SchCD) = "B" Then
            'lbl_2Sel(0).Caption = ""
            'lbl_2Sel(1).Caption = ""
        End If
        
        
    End With
    
End Sub


Private Function Div_Gwamok_NM(ByVal aGbn As String, ByVal aGwamok As String) As String
    Dim sDiv()      As String
    Dim ni          As Integer
    Dim sTmp        As String
    
    Dim sGwamok     As String
    
    sDiv = Split(aGwamok, "|", -1, vbTextCompare)
    
    sTmp = "":  sGwamok = ""
    For ni = 0 To UBound(sDiv) - 1 Step 1
        
        sTmp = sDiv(ni)
        
        Select Case aGbn
            Case "SEL1"
                Select Case sTmp
                    Case constSatamCodes(0)
                        sTmp = constSatams(0)
                    Case constSatamCodes(1)
                        sTmp = constSatams(1)
                    Case constSatamCodes(2)
                        sTmp = constSatams(2)
                    Case constSatamCodes(3)
                        sTmp = constSatams(3)
                    Case constSatamCodes(4)
                        sTmp = constSatams(4)
                    Case constSatamCodes(5)
                        sTmp = constSatams(5)
                    Case constSatamCodes(6)
                        sTmp = constSatams(6)
                    Case constSatamCodes(7)
                        sTmp = constSatams(7)
                    Case constSatamCodes(8)
                        sTmp = constSatams(8)
                    Case constSatamCodes(9)
                        sTmp = constSatams(9)
'                    Case "11"
'                        sTmp = "세계지리"
                End Select
            Case "SEL2"
                Select Case sTmp
                    Case "31"
                        sTmp = "독어"
                    Case "32"
                        sTmp = "일어"
                    Case "33"
                        sTmp = "에스파냐어"
                    Case "34"
                        sTmp = "불어"
                    Case "35"
                        sTmp = "중국어"
                    Case "36"
                        sTmp = "한문"
                End Select
            Case "SEL3"
                Select Case sTmp
                    Case "51"
                        sTmp = "물리1"
                    Case "52"
                        sTmp = "화학1"
                    Case "53"
                        sTmp = "생명과학1"
                    Case "54"
                        sTmp = "지구과학1"
                    Case "55"
                        sTmp = "물리2"
                    Case "56"
                        sTmp = "화학2"
                    Case "57"
                        sTmp = "생명과학2"
                    Case "58"
                        sTmp = "지구과학2"
                End Select
            Case "SEL4"
                Select Case sTmp
                    Case "81"
                        sTmp = "미적분"
                    Case "82"
                        sTmp = "이산수학"
                    Case "83"
                        sTmp = "확률통계"
                    Case "84"
                        sTmp = "수리나형"
                End Select
            Case "SEL5"
                Select Case sTmp
                    Case "91"
                        sTmp = "언어"
                    Case "92"
                        sTmp = "수리"
                    Case "93"
                        sTmp = "외국어"             '< 변경
                    Case "94"
                        sTmp = ""                   '< 변경
                End Select
            Case Else
                sTmp = ""
        End Select
            
        If ni > 0 Then sGwamok = sGwamok & ", "
        sGwamok = sGwamok & sTmp
        
    Next ni
    
    If sGwamok = "" Then
        Div_Gwamok_NM = ""
    Else
        Div_Gwamok_NM = sGwamok
    End If
    
End Function

'>> 이미지 받은파일 체크 : 체크시 이상이 있는 경우엔 default 값을 보여줌.
Public Function CheckJPG(fileName As String) As Picture

    Dim header(2)     As Byte
    Dim tailer(2)     As Byte
    Dim f             As Integer
    Dim MaxSize       As Long


    On Error Resume Next

    f = FreeFile()
    Open fileName For Binary As #f

        On Error GoTo 0
        If Err <> 0 Then
            Set CheckJPG = imgList.ListImages.Item(1).Picture
            Exit Function
        End If

        On Error Resume Next
        MaxSize = LOF(f)                                        '<< 파일의 바이트 크기를 구합니다.
        Get #f, 1, header()
        Get #f, MaxSize - 1, tailer()
    Close f

    ' Must start with hex FF D8  and end data hex FF D9
    If (header(0) = 255 And header(1) = 216) And _
       (tailer(0) = 255 And tailer(1) >= 209) Then
        Set CheckJPG = LoadPicture(fileName)
    Else
        Set CheckJPG = imgList.ListImages.Item(1).Picture       '<< no-image
    End If

End Function


'## 서버의 이미지 가져오기
Private Sub Get_STD_image()
    
    Dim bData()     As Byte
    Dim f           As Integer
    Dim nRec        As Long

    Dim sLocalFile  As String
    Dim sSourceUrl  As String
    
    Dim MaxSize     As Long

    On Error Resume Next

    f = FreeFile()
    
    For nRec = 1 To UBound(uSTD) Step 1
    
        sLocalFile = sSavePath & "\" & uSTD(nRec).IMAGE_FILE & ".jpg"               '<< unique key : 학원+수험번호
        
        If Dir(sLocalFile, vbNormal) > " " Then
            Open sLocalFile For Binary As #f
                On Error Resume Next
                MaxSize = LOF(f)
            Close f
            
            If MaxSize = 0 Then
                Kill sLocalFile
            End If
        End If
        
        If Dir(sLocalFile, vbNormal) = "" Then                                      '<< 학생 이미지 없는 것만 받음
        
            Select Case Trim(basModule.SchCD)
                Case "B"        '<< 부산직영
                    sSourceUrl = "http://www.dsbusan.com" & uSTD(nRec).PHOTO_PATH           '<< 서버의 이미지 경로
                Case Else
                    sSourceUrl = "http://www.dshw.co.kr" & uSTD(nRec).PHOTO_PATH            '<< 서버의 이미지 경로
            End Select
             
            bData() = Inet1.OpenURL(sSourceUrl, icByteArray)
            
            If UBound(bData) > 0 Then
                Open sLocalFile For Binary Access Write As #f
                Put #f, , bData()
            
                DoEvents
                Close #f
            End If
            
        End If
    Next nRec
    
End Sub


'## 전체 출력

Private Sub cmdPrintAll_Click()

    Dim nRec        As Long
    Dim bChk        As Boolean

    If UBound(uSTD) < 1 Then
        Exit Sub
    End If
    
    bChk = False
    With dlgPrint
        .CancelError = True
        .ShowPrinter
        
        bChk = True
    End With
    
ErrPrint:
    If bChk = False Then
        MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "현재페이지 인쇄하기"
        Exit Sub
    End If
    
    nRec = 0
    cmdPrint.Tag = "ALL"
    
    Do
        nRec = nRec + 1
        txtPage.Text = Trim(CStr(nRec)) & "/" & Trim(CStr(UBound(uSTD)))
        
        
        Call Std_Data_Show(nRec)                                '<< 학생자료 화면 보이기
        Me.Tag = "LOAD"
            VScroll1.value = nRec
            Call CmdPrint_Click:        DoEvents                '<< 1명 출력
            
        Me.Tag = ""

    Loop Until nRec = UBound(uSTD)
    
    cmdPrint.Tag = ""
    MsgBox "출력을 완료하였습니다.", vbInformation + vbOKOnly, "전체출력"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    cmdPrint.Tag = ""
    
    MsgBox "출력시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체출력"
    
End Sub

'## 현재 페이지 출력 : 1명 출력
Public Sub CmdPrint_Click()

    Dim i           As Integer
    Dim X           As Integer
    Dim Y           As Integer
    Dim pRate       As Double


    Dim bChk        As Boolean

    If UBound(uSTD) < 1 Then
        Exit Sub
    End If

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
            MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "현재페이지 인쇄하기"
            Exit Sub
        End If
    End If
    
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
                 Printer.DrawWidth = 1                   ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent     ' 단색
                 Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And _
                 UCase(UsrCtl.Name) = "FILLBOXS") Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                 Printer.DrawWidth = 1                   ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent     ' 단색
                 Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
             
             
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And _
                 (UCase(UsrCtl.Name) = "L_B3" _
                 )) Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                
                Select Case Trim(basModule.SchCD)
                    Case "B"
                    
                    Case Else
                        Printer.DrawWidth = 1                   ' 선의 굵기
                        Printer.FillStyle = vbFSTransparent     ' 단색
                        Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                        PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
                 End Select
                 
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
             
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And _
                 (UCase(UsrCtl.Name) = "L_B2" _
                 )) Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                
                Select Case Trim(basModule.SchCD)
                    Case "B"
                    
                    Case Else
                        Printer.DrawWidth = 12
                        PrintBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
                 End Select
                 
             End If
             
             
        End With
    Next

    '김한욱 시작
    For Each UsrCtl In Me
        With UsrCtl
             Select Case UCase(TypeName(UsrCtl))
                    Case "LINE"

                        If Trim(basModule.SchCD) = "B" Then
                            If UCase(UsrCtl.Name) = "L_B1" Then
                                
                            Else
                                Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                                Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                                Printer.FillStyle = vbFSTransparent
                                PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate
                            End If
                         Else
                                Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                                Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                                Printer.FillStyle = vbFSTransparent
                                PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate
                         End If
                    Case "LABEL"
                          '********************************************************************
                          '  Label을 그대로 출력 한다(속성)
                          '  단) transparent는 true로 처리하고 실행한다.
                          '  SetBkMode(Printer.hdc, TRANSPARENT)문장은 MS버그를 처리하기 위함
                          '********************************************************************
                            If Trim(basModule.SchCD) = "B" Then
                                If (UCase(UsrCtl.Name) = "B_LB") Then
                                Else
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
                                End If
                            Else
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
                            End If
                    Case "TEXTBOX"
                         '********************************************************************
                         '  데이터 출력 (DATA는 TEXTBOX로 처리 한다.)
                         '********************************************************************
                          Select Case UCase(.Name)
                            Case "TXTSTDNM", "TXTPAGE"
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
                          '  사진출력
                          '********************************************************************
                          If (Photo.Picture <> 0) Then
                              Printer.FontTransparent = True
                              iBKMode = SetBkMode(Printer.hDC, OPAQUE)
                              ' iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
                              PrintPicture .Picture, .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
                          End If
             End Select
        End With
    Next

          
        
        For Each UsrCtl In Me
        With UsrCtl

             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS1") Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                 Printer.DrawWidth = 1                   ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent     ' 단색
                 Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
        End With
        Next
    
        For Each UsrCtl In Me
            With UsrCtl
                 If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS1") Then
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
                 Select Case UCase(UsrCtl.Name)
                        Case "Lines1"
                             '********************************************************************
                             '  박스/line를 긋는다.
                             '********************************************************************
                              Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                              Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                              Printer.FillStyle = vbFSTransparent
                              PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate
    
                        
                 End Select
            End With
        Next '김한욱 끝
    
        Printer.EndDoc     ' 프린터로 보낸다

End Sub



'## 사진 업로드
Private Sub Photo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sFileLocation   As String
    Dim sSchNO          As String
    Dim sOrdNO          As String
    
    Dim sExmID          As String
    Dim simageFile      As String

    Dim bRet            As String
    
    Dim sDiv()          As String
    Dim nS              As Long
    Dim sLocalFile      As String
    
    If Button <> vbRightButton Then
        Exit Sub
    End If

    If 학생성명.Tag = "" Then
        MsgBox "학생을 조회하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
        Exit Sub
    End If
    If UBound(uSTD) < 1 Then
        MsgBox "학생을 조회하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
        Exit Sub
    End If
    
    '수험번호.tag
    
    With uSTD(VScroll1.value)
        sOrdNO = .ORD_NO
        sSchNO = .SCHNO
        sFileLocation = .IMAGE_DIR
        
        
        bRet = ""
        If Trim(sOrdNO) = "" Then        '< 이미지가 없는 경우엔 강제로 생성
            bRet = Make_image_Path(sSchNO, sExmID, simageFile)
            
            If bRet = "" Then
                MsgBox "경로 생성에 문제가 있습니다." & vbCrLf & _
                       "관리자에게 문의하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
                Exit Sub
            Else
                sFileLocation = bRet
            End If
        End If
        
    End With
    
    
    '<< 파일 지우기 >>
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        sLocalFile = sSavePath & "\" & uSTD(nS).IMAGE_FILE & ".jpg"       '<< unique key : 학원+수험번호
        If Dir(sLocalFile) > " " Then
            Kill sLocalFile
        End If
    End If
    
    Load INT900
    Call INT900.Save_Photo(sFileLocation, sSchNO)
    INT900.Show
    
End Sub



'## 이미지 없는 경우 강제를 생성
Private Function Make_image_Path(ByVal aSchNO As String, ByVal aExmID As String, ByVal aimageFile As String) As String
    Dim sFilePath       As String
    
    Dim sStr            As String
    Dim DBCmd           As ADODB.Command
    Dim DBParam         As ADODB.Parameter
    
    Dim ni              As Long
    Dim sLocalFile      As String
    Dim nExe            As Integer
    Dim f               As Integer
    Dim MaxSize         As Long
    
    sFilePath = ""
    Select Case Trim(basModule.SchCD)
        Case "N"
            sFilePath = "/NDOC/dshw/noryangjin/register/ACC/"
        Case "K", "W", "Q"
            sFilePath = "/NDOC/dshw/kangnam/register/ACC/"
        Case "S"
            sFilePath = "/NDOC/dshw/songpa/register/ACC/"
        Case "P"
            sFilePath = "/NDOC/dshw/msongpa/register/ACC/"
        Case "M"
            sFilePath = "/NDOC/dshw/mkangnam/register/ACC/"
        Case "J"
            sFilePath = "/NDOC/dshw/mgwanghwa/register/ACC/"
        Case "B"
            sFilePath = "/NDOC/dshw/busan/register/ACC/"
    End Select
    
    sFilePath = sFilePath & Trim(basModule.SchCD) & aExmID & ".jpg"
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    '<< UPDATE
    sStr = ""
    sStr = sStr & " Update CLSTD01TB"
    sStr = sStr & "    SET PHOTO_PATH = '" & sFilePath & "'"
    sStr = sStr & "  WHERE SCHNO = '" & Trim(aSchNO) & "'"
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        f = FreeFile()
        sLocalFile = sSavePath & "\" & aimageFile & ".jpg"               '<< unique key : 학원+수험번호
        If Dir(sLocalFile) > " " Then
            Open sLocalFile For Binary As #f
                On Error Resume Next
                MaxSize = LOF(f)
            Close f
            
            Kill sLocalFile
            
        End If
    
        Make_image_Path = sFilePath
    Else
        basDataBase.DBConn.RollbackTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
    
        Make_image_Path = ""
    End If
    
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Make_image_Path = ""
End Function




