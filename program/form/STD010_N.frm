VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form STD010_N 
   Caption         =   "입학사정 >> 학생등록(노량진,송파)"
   ClientHeight    =   12960
   ClientLeft      =   4545
   ClientTop       =   2790
   ClientWidth     =   18510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   17926.62
   ScaleMode       =   0  '사용자
   ScaleWidth      =   19402.51
   Begin VB.Frame fraClinic 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame17"
      Height          =   3105
      Left            =   8460
      TabIndex        =   210
      Top             =   10200
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame frame 
         BackColor       =   &H00C6AD84&
         BorderStyle     =   0  '없음
         Caption         =   ">> 논술 선택과목"
         Height          =   3045
         Index           =   2
         Left            =   30
         TabIndex        =   211
         Top             =   60
         Width           =   9075
         Begin VB.CommandButton cmdClinicClear 
            Caption         =   "전체 선택 해제"
            Height          =   495
            Left            =   4620
            TabIndex        =   230
            Top             =   2520
            Width           =   1425
         End
         Begin VB.CommandButton cmdClinicOK 
            Caption         =   "확인"
            Height          =   495
            Left            =   2760
            TabIndex        =   229
            Top             =   2520
            Width           =   1425
         End
         Begin VB.Frame Frame25 
            BackColor       =   &H00F7EFE7&
            Caption         =   "수학"
            Height          =   675
            Left            =   240
            TabIndex        =   217
            Top             =   1050
            Width           =   8715
            Begin VB.OptionButton chkClinic_M 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_M(0)"
               Height          =   375
               Index           =   0
               Left            =   210
               TabIndex        =   224
               Top             =   210
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_M 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_M(2)"
               Height          =   375
               Index           =   2
               Left            =   4380
               TabIndex        =   223
               Top             =   210
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_M 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_M(3)"
               Height          =   375
               Index           =   3
               Left            =   6450
               TabIndex        =   222
               Top             =   150
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_M 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_M(1)"
               Height          =   375
               Index           =   1
               Left            =   2340
               TabIndex        =   221
               Top             =   210
               Visible         =   0   'False
               Width           =   2085
            End
         End
         Begin VB.Frame Frame24 
            BackColor       =   &H00F7EFE7&
            Caption         =   "영어"
            Height          =   675
            Left            =   240
            TabIndex        =   216
            Top             =   1770
            Width           =   8715
            Begin VB.OptionButton chkClinic_E 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_E(0)"
               Height          =   375
               Index           =   0
               Left            =   210
               TabIndex        =   228
               Top             =   240
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_E 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_E(2)"
               Height          =   375
               Index           =   2
               Left            =   4380
               TabIndex        =   227
               Top             =   180
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_E 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_E(3)"
               Height          =   375
               Index           =   3
               Left            =   6450
               TabIndex        =   226
               Top             =   180
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_E 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_E(1)"
               Height          =   375
               Index           =   1
               Left            =   2310
               TabIndex        =   225
               Top             =   210
               Visible         =   0   'False
               Width           =   2085
            End
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H00F7EFE7&
            Caption         =   "국어"
            Height          =   675
            Left            =   240
            TabIndex        =   214
            Top             =   360
            Width           =   8715
            Begin VB.OptionButton chkClinic_L 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_L(2)"
               Height          =   435
               Index           =   2
               Left            =   4380
               TabIndex        =   220
               Top             =   180
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_L 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_L(3)"
               Height          =   435
               Index           =   3
               Left            =   6450
               TabIndex        =   219
               Top             =   180
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_L 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_L(1)"
               Height          =   435
               Index           =   1
               Left            =   2340
               TabIndex        =   218
               Top             =   180
               Visible         =   0   'False
               Width           =   2085
            End
            Begin VB.OptionButton chkClinic_L 
               BackColor       =   &H00F7EFE7&
               Caption         =   "chkClinic_L(0)"
               Height          =   435
               Index           =   0
               Left            =   210
               TabIndex        =   215
               Top             =   180
               Width           =   2085
            End
         End
         Begin VB.Label Label15 
            BackStyle       =   0  '투명
            Caption         =   ">> 클리닉 선택"
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
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   212
            Top             =   90
            Width           =   2625
         End
      End
   End
   Begin VB.CheckBox chkSatam 
      BackColor       =   &H00F7EFE7&
      Caption         =   "미선택"
      Height          =   345
      Index           =   12
      Left            =   7110
      TabIndex        =   198
      Top             =   4980
      Visible         =   0   'False
      Width           =   1188
   End
   Begin VB.Frame FraPay 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   2295
      Left            =   15360
      TabIndex        =   168
      Top             =   540
      Width           =   5625
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   2235
         Left            =   90
         TabIndex        =   169
         Top             =   150
         Width           =   5565
         Begin VB.TextBox txtPayChg 
            Height          =   270
            IMEMode         =   10  '한글 
            Left            =   1470
            TabIndex        =   180
            Text            =   "txtPayChg"
            Top             =   810
            Width           =   1605
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '없음
            Height          =   435
            Left            =   1260
            TabIndex        =   178
            Top             =   1140
            Width           =   3405
            Begin VB.OptionButton OptPay1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "카드"
               Height          =   285
               Left            =   0
               TabIndex        =   171
               Top             =   90
               Width           =   885
            End
            Begin VB.OptionButton OptPay2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "핸드폰"
               Height          =   285
               Left            =   1830
               TabIndex        =   172
               Top             =   90
               Width           =   885
            End
         End
         Begin VB.TextBox txtPay 
            Height          =   270
            IMEMode         =   10  '한글 
            Left            =   1260
            TabIndex        =   170
            Text            =   "txtPay"
            Top             =   240
            Width           =   1605
         End
         Begin VB.ComboBox cboCard 
            Height          =   300
            Left            =   1230
            Style           =   2  '드롭다운 목록
            TabIndex        =   173
            Top             =   1560
            Width           =   1725
         End
         Begin VB.CommandButton cmdPaySave 
            Caption         =   "등록하기"
            Height          =   450
            Left            =   3840
            TabIndex        =   174
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label43 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문번호가 변경시에만"
            Height          =   210
            Left            =   900
            TabIndex        =   181
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label57 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주문번호"
            Height          =   210
            Left            =   -150
            TabIndex        =   177
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label55 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "결제"
            Height          =   210
            Left            =   -150
            TabIndex        =   176
            Top             =   1290
            Width           =   1185
         End
         Begin VB.Label Label58 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4470
            TabIndex        =   175
            Top             =   150
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraPoint 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   5295
      Left            =   60
      TabIndex        =   161
      Top             =   12030
      Width           =   7185
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   5235
         Left            =   60
         TabIndex        =   162
         Top             =   30
         Width           =   7125
         Begin VB.CommandButton cmdAddPointRow 
            Caption         =   "과목점수 추가"
            Height          =   450
            Left            =   930
            TabIndex        =   166
            Top             =   4680
            Width           =   1635
         End
         Begin VB.CommandButton cmdSavePoint 
            Caption         =   "점수등록"
            Height          =   450
            Left            =   4380
            TabIndex        =   163
            Top             =   4680
            Width           =   2595
         End
         Begin FPSpread.vaSpread sprPoint 
            Height          =   4035
            Left            =   30
            TabIndex        =   165
            Top             =   510
            Width           =   7035
            _Version        =   393216
            _ExtentX        =   12409
            _ExtentY        =   7117
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
            MaxCols         =   9
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD010_N.frx":0000
         End
         Begin VB.Label Label54 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "과목을 신규등록내용 삭제시 삭제내용 선택후 del 키를 누르세요."
            Height          =   210
            Left            =   -270
            TabIndex        =   167
            Top             =   270
            Width           =   5625
         End
         Begin VB.Label Label56 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   6030
            TabIndex        =   164
            Top             =   120
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraAddr 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '없음
      Height          =   3015
      Left            =   15300
      TabIndex        =   150
      Top             =   5280
      Width           =   6315
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   2955
         Left            =   30
         TabIndex        =   151
         Top             =   30
         Width           =   6255
         Begin VB.TextBox txtEmail 
            Height          =   345
            Left            =   1170
            TabIndex        =   80
            Text            =   "txtEmail"
            Top             =   1800
            Width           =   4605
         End
         Begin VB.CommandButton cmdSaveAddr 
            Caption         =   "확인"
            Height          =   450
            Left            =   4290
            TabIndex        =   81
            Top             =   2310
            Width           =   1395
         End
         Begin VB.TextBox txtAddr2 
            Height          =   345
            Left            =   1170
            TabIndex        =   79
            Text            =   "txtAddr2"
            Top             =   1380
            Width           =   4605
         End
         Begin VB.TextBox txtAddr1 
            Height          =   345
            Left            =   1170
            TabIndex        =   78
            Text            =   "txtAddr1"
            Top             =   990
            Width           =   4605
         End
         Begin EditLib.fpMask fpZip 
            Height          =   345
            Left            =   1170
            TabIndex        =   77
            Top             =   570
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
            Mask            =   "###-###"
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
         Begin EditLib.fpMask fpBirth_ymdS 
            Height          =   345
            Left            =   1170
            TabIndex        =   76
            Top             =   150
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
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
            Mask            =   "9999-99-99"
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
         Begin VB.Label Label51 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "이메일"
            Height          =   210
            Left            =   90
            TabIndex        =   157
            Top             =   1867
            Width           =   975
         End
         Begin VB.Label Label50 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   90
            TabIndex        =   156
            Top             =   217
            Width           =   975
         End
         Begin VB.Label Label49 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   5070
            TabIndex        =   155
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label48 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "주소"
            Height          =   210
            Left            =   90
            TabIndex        =   154
            Top             =   1447
            Width           =   975
         End
         Begin VB.Label Label47 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "우편주소"
            Height          =   210
            Left            =   90
            TabIndex        =   153
            Top             =   1057
            Width           =   975
         End
         Begin VB.Label Label46 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "우편번호"
            Height          =   210
            Left            =   90
            TabIndex        =   152
            Top             =   637
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraGwamok 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "과목"
      Height          =   4275
      Left            =   7380
      TabIndex        =   131
      Top             =   12480
      Width           =   8865
      Begin VB.Frame Frame23 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   4215
         Left            =   30
         TabIndex        =   132
         Top             =   30
         Width           =   8805
         Begin VB.CommandButton cmdClose 
            Caption         =   "닫기"
            Height          =   330
            Left            =   8160
            TabIndex        =   133
            Top             =   3840
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   4080
            Left            =   30
            Picture         =   "STD010_N.frx":03F1
            Top             =   60
            Width           =   8730
         End
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame20"
      Height          =   4300
      Left            =   8460
      TabIndex        =   127
      Top             =   6150
      Width           =   6615
      Begin VB.Frame Frame21 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame21"
         Height          =   4240
         Left            =   30
         TabIndex        =   128
         Top             =   30
         Width           =   6555
         Begin VB.CommandButton cmdGwamokView 
            Caption         =   "과목보기"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   82
            Top             =   870
            Width           =   885
         End
         Begin VB.CommandButton cmdExcelSave 
            Caption         =   "엑셀자료 저장하기"
            Enabled         =   0   'False
            Height          =   450
            Left            =   4590
            TabIndex        =   74
            Top             =   3765
            Width           =   1875
         End
         Begin VB.CommandButton cmdGetExcel 
            Caption         =   "엑셀자료 가져오기"
            Enabled         =   0   'False
            Height          =   390
            Left            =   4410
            TabIndex        =   73
            Top             =   90
            Width           =   1875
         End
         Begin MSComDlg.CommonDialog dlgFile 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin FPSpread.vaSpread sprExcel_STD_Data 
            Height          =   2445
            Left            =   60
            TabIndex        =   83
            Top             =   1230
            Width           =   6405
            _Version        =   393216
            _ExtentX        =   11298
            _ExtentY        =   4313
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
            MaxCols         =   16
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD010_N.frx":7ABB
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '투명
            Caption         =   $"STD010_N.frx":8011
            Height          =   615
            Left            =   240
            TabIndex        =   130
            Top             =   630
            Width           =   5385
         End
         Begin VB.Label Label29 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회기본항목"
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
            Left            =   120
            TabIndex        =   129
            Top             =   120
            Width           =   2625
         End
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame18"
      Height          =   6045
      Left            =   8460
      TabIndex        =   118
      Top             =   60
      Width           =   6615
      Begin VB.Frame Frame19 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame19"
         Height          =   5985
         Left            =   30
         TabIndex        =   119
         Top             =   30
         Width           =   6555
         Begin VB.TextBox Text1 
            Height          =   2175
            Left            =   1080
            TabIndex        =   204
            Text            =   "Text1"
            Top             =   2880
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton cmdAllStdData 
            Caption         =   "엑셀로 데이터 받기"
            Height          =   315
            Left            =   1740
            TabIndex        =   57
            Top             =   30
            Width           =   2955
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   5220
            Style           =   2  '드롭다운 목록
            TabIndex        =   61
            Top             =   450
            Width           =   885
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   810
            Style           =   2  '드롭다운 목록
            TabIndex        =   62
            Top             =   780
            Width           =   855
         End
         Begin EditLib.fpLongInteger fpPayOK 
            Height          =   315
            Left            =   3480
            TabIndex        =   64
            Top             =   765
            Width           =   645
            _Version        =   196608
            _ExtentX        =   1138
            _ExtentY        =   556
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
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
         Begin VB.ComboBox cboPay 
            Height          =   300
            Left            =   5730
            Style           =   2  '드롭다운 목록
            TabIndex        =   66
            Top             =   765
            Width           =   855
         End
         Begin VB.ComboBox cboPassCN 
            Height          =   300
            Left            =   4710
            Style           =   2  '드롭다운 목록
            TabIndex        =   69
            Top             =   1140
            Width           =   885
         End
         Begin VB.ComboBox cboKaeyol_F 
            Height          =   300
            Left            =   3210
            Style           =   2  '드롭다운 목록
            TabIndex        =   60
            Top             =   420
            Width           =   915
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회하기(&F)"
            Height          =   390
            Left            =   4530
            TabIndex        =   72
            Top             =   1470
            Width           =   1305
         End
         Begin VB.TextBox txtStdNM_F 
            Height          =   345
            IMEMode         =   10  '한글 
            Left            =   810
            TabIndex        =   67
            Text            =   "txtStdNM_F"
            Top             =   1125
            Width           =   825
         End
         Begin VB.ComboBox cboSel1_SCH_F 
            Height          =   300
            Left            =   810
            Style           =   2  '드롭다운 목록
            TabIndex        =   70
            Top             =   1515
            Width           =   1005
         End
         Begin VB.ComboBox cboSel2_SCH_F 
            Height          =   300
            Left            =   2790
            Style           =   2  '드롭다운 목록
            TabIndex        =   71
            Top             =   1515
            Width           =   1275
         End
         Begin EditLib.fpMask fpExmID_F 
            Height          =   345
            Left            =   810
            TabIndex        =   58
            Top             =   390
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
         Begin EditLib.fpMask fpBirth_ymd_F 
            Height          =   345
            Left            =   2430
            TabIndex        =   68
            Top             =   1125
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
            Mask            =   "9999-99-99"
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
            Height          =   345
            Left            =   1920
            TabIndex        =   59
            Top             =   390
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
         Begin EditLib.fpLongInteger fpPayNot 
            Height          =   315
            Left            =   4710
            TabIndex        =   65
            Top             =   765
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
            _ExtentY        =   556
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
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
         Begin EditLib.fpLongInteger fpPayTot 
            Height          =   315
            Left            =   2430
            TabIndex        =   63
            Top             =   765
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   556
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
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483648"
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
         Begin FPSpread.vaSpread sprSTD_F 
            Height          =   4035
            Left            =   30
            TabIndex        =   149
            Top             =   1890
            Width           =   6495
            _Version        =   393216
            _ExtentX        =   11456
            _ExtentY        =   7117
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
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
            MaxCols         =   40
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            SpreadDesigner  =   "STD010_N.frx":80A8
         End
         Begin VB.Image imgExcel 
            Height          =   420
            Left            =   6150
            Picture         =   "STD010_N.frx":8A34
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label Label38 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "전체결재"
            ForeColor       =   &H00C000C0&
            Height          =   210
            Left            =   1440
            TabIndex        =   141
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label37 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "인터넷/학원"
            Height          =   210
            Left            =   4110
            TabIndex        =   140
            Top             =   495
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "무/유시험"
            Height          =   210
            Left            =   -150
            TabIndex        =   139
            Top             =   825
            Width           =   975
         End
         Begin VB.Label Label35 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "미결재"
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   3720
            TabIndex        =   138
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label34 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "결재"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3030
            TabIndex        =   137
            Top             =   810
            Width           =   435
         End
         Begin VB.Label Label33 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "결재"
            Height          =   210
            Left            =   5250
            TabIndex        =   136
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label32 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합격차수"
            Height          =   210
            Left            =   3720
            TabIndex        =   135
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계  열"
            Height          =   210
            Left            =   2160
            TabIndex        =   134
            Top             =   465
            Width           =   1035
         End
         Begin VB.Label Label27 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   1440
            TabIndex        =   125
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            Height          =   210
            Left            =   -240
            TabIndex        =   124
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  '투명
            Caption         =   "수험번호             부터"
            Height          =   210
            Left            =   30
            TabIndex        =   123
            Top             =   450
            Width           =   2025
         End
         Begin VB.Label Label24 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회기본항목"
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
            Left            =   120
            TabIndex        =   122
            Top             =   90
            Width           =   2625
         End
         Begin VB.Label Label23 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "1지망학원"
            Height          =   210
            Left            =   -180
            TabIndex        =   121
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label22 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "2지망학원"
            Height          =   210
            Left            =   1770
            TabIndex        =   120
            Top             =   1560
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame10"
      Height          =   11115
      Left            =   60
      TabIndex        =   84
      Top             =   30
      Width           =   8355
      Begin VB.Frame Frame9 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame9"
         Height          =   11010
         Left            =   30
         TabIndex        =   85
         Top             =   30
         Width           =   8295
         Begin VB.CommandButton cmdCancel 
            Caption         =   "학생 취소하기"
            Height          =   450
            Left            =   4410
            TabIndex        =   51
            Top             =   10095
            Width           =   1815
         End
         Begin VB.CommandButton cmdStdDel 
            Caption         =   "학생삭제하기"
            Height          =   450
            Left            =   6690
            TabIndex        =   52
            Top             =   10095
            Width           =   1365
         End
         Begin VB.CommandButton cmdStdin 
            Caption         =   "학생등록 및 수정하기 (&S)"
            Height          =   450
            Left            =   900
            TabIndex        =   50
            Top             =   10095
            Width           =   2655
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '없음
            Caption         =   "Frame17"
            Height          =   825
            Index           =   0
            Left            =   30
            TabIndex        =   109
            Top             =   9135
            Width           =   8235
            Begin VB.Frame fraSEL5 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 논술 선택과목"
               Height          =   765
               Index           =   0
               Left            =   30
               TabIndex        =   110
               Top             =   30
               Width           =   8175
               Begin VB.CommandButton cmdSelClinic 
                  Caption         =   ">> 클리닉 "
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   213
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "언어"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   46
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "수리"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   47
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "외국어"
                  Height          =   345
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   48
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "과학탐구"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   49
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 탐구선택"
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
                  Height          =   285
                  Index           =   0
                  Left            =   90
                  TabIndex        =   111
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '없음
            Caption         =   "Frame16"
            Height          =   825
            Left            =   30
            TabIndex        =   106
            Top             =   8265
            Width           =   8235
            Begin VB.Frame fraSEL4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 수리영역 선택과목"
               Height          =   765
               Left            =   30
               TabIndex        =   107
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "미적분"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   42
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "이산수학"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   43
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "확률통계"
                  Height          =   345
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   44
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "수리나형"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   45
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.Label Label14 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 수리영역 선택과목"
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
                  Height          =   285
                  Left            =   90
                  TabIndex        =   108
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '없음
            Caption         =   "Frame15"
            Height          =   1215
            Left            =   30
            TabIndex        =   103
            Top             =   7005
            Width           =   8235
            Begin VB.Frame fraSEL3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 과학탐구 선택과목"
               Height          =   1155
               Left            =   30
               TabIndex        =   104
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "특강"
                  Height          =   345
                  Index           =   9
                  Left            =   5580
                  TabIndex        =   235
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "물리1"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   34
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "화학1"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   35
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "생명과학1"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   36
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "지구과학1"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   37
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "물리2"
                  Height          =   345
                  Index           =   5
                  Left            =   240
                  TabIndex        =   38
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "화학2"
                  Height          =   345
                  Index           =   6
                  Left            =   1620
                  TabIndex        =   39
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "생명과학2"
                  Height          =   345
                  Index           =   7
                  Left            =   2970
                  TabIndex        =   40
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "지구과학2"
                  Height          =   345
                  Index           =   8
                  Left            =   4320
                  TabIndex        =   41
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.Label Label13 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 과학탐구 선택과목"
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
                  Height          =   285
                  Left            =   90
                  TabIndex        =   105
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '없음
            Caption         =   "Frame14"
            Height          =   825
            Left            =   30
            TabIndex        =   100
            Top             =   6135
            Width           =   8235
            Begin VB.Frame fraSEL2 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 제2 외국어 선택과목"
               Height          =   765
               Left            =   30
               TabIndex        =   101
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "베트남어"
                  Height          =   345
                  Index           =   13
                  Left            =   7140
                  TabIndex        =   236
                  Top             =   -30
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "아랍어"
                  Height          =   345
                  Index           =   12
                  Left            =   7140
                  TabIndex        =   33
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "세계지리"
                  Height          =   345
                  Index           =   11
                  Left            =   5820
                  TabIndex        =   32
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "세계사"
                  Height          =   345
                  Index           =   10
                  Left            =   4320
                  TabIndex        =   31
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "외국어"
                  Height          =   345
                  Index           =   9
                  Left            =   2970
                  TabIndex        =   30
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "수리"
                  Height          =   345
                  Index           =   8
                  Left            =   1620
                  TabIndex        =   29
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "언어"
                  Height          =   345
                  Index           =   7
                  Left            =   240
                  TabIndex        =   28
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "독어"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "일어"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   23
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "에스파냐어"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   24
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "불어"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   25
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "중국어"
                  Height          =   345
                  Index           =   5
                  Left            =   5820
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "한문"
                  Height          =   345
                  Index           =   6
                  Left            =   7140
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 제2 외국어 선택과목"
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
                  Height          =   285
                  Left            =   90
                  TabIndex        =   102
                  Top             =   60
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '없음
            Caption         =   "Frame13"
            Height          =   1215
            Left            =   30
            TabIndex        =   99
            Top             =   4875
            Width           =   8235
            Begin VB.Frame fraSEL1 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 사회탐구 선택과목"
               Height          =   1155
               Left            =   30
               TabIndex        =   186
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "특강"
                  Height          =   345
                  Index           =   11
                  Left            =   6960
                  TabIndex        =   234
                  Top             =   750
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "한국사"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   196
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "세계사"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   195
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "동아시아사"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   194
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "한국지리"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   193
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "세계지리"
                  Height          =   345
                  Index           =   5
                  Left            =   5790
                  TabIndex        =   192
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "생활과윤리"
                  Height          =   345
                  Index           =   6
                  Left            =   240
                  TabIndex        =   191
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "윤리와사상"
                  Height          =   345
                  Index           =   7
                  Left            =   1620
                  TabIndex        =   190
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "법과정치"
                  Height          =   345
                  Index           =   8
                  Left            =   2970
                  TabIndex        =   189
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "경제"
                  Height          =   345
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   188
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "사회문화"
                  Height          =   345
                  Index           =   10
                  Left            =   5790
                  TabIndex        =   187
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 사회탐구 선택과목"
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
                  Height          =   285
                  Left            =   60
                  TabIndex        =   197
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '없음
            Caption         =   "Frame12"
            Height          =   1200
            Left            =   30
            TabIndex        =   94
            Top             =   3645
            Width           =   8235
            Begin VB.Frame Frame4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 점수"
               Height          =   1150
               Left            =   30
               TabIndex        =   95
               Top             =   30
               Width           =   8175
               Begin VB.ComboBox cbo_Tamgoo 
                  Height          =   300
                  Left            =   660
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   202
                  Top             =   750
                  Width           =   1215
               End
               Begin VB.CommandButton cmdAddPoint 
                  Caption         =   "학생상세점수(&P)"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4440
                  TabIndex        =   160
                  Top             =   720
                  Width           =   1725
               End
               Begin VB.CommandButton cmdChgAddr 
                  Caption         =   "학생상세변경"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   6360
                  TabIndex        =   75
                  Top             =   720
                  Width           =   1665
               End
               Begin VB.TextBox txtCancel 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   6390
                  TabIndex        =   17
                  Text            =   "txtCancel"
                  Top             =   0
                  Width           =   1695
               End
               Begin EditLib.fpLongInteger fpK_Num 
                  Height          =   345
                  Left            =   600
                  TabIndex        =   18
                  Top             =   300
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
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
                  UserEntry       =   1
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
                  MaxValue        =   "9999"
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
               Begin EditLib.fpLongInteger fpE_Num 
                  Height          =   345
                  Left            =   3480
                  TabIndex        =   20
                  Top             =   300
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
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
                  UserEntry       =   1
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
                  MaxValue        =   "9999"
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
               Begin EditLib.fpLongInteger fpM_Num 
                  Height          =   345
                  Left            =   1920
                  TabIndex        =   19
                  Top             =   300
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
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
                  UserEntry       =   1
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
                  MaxValue        =   "9999"
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
               Begin EditLib.fpLongInteger fpT_Num 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   203
                  Top             =   720
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
                  _ExtentY        =   556
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
                  MaxValue        =   "2147483647"
                  MinValue        =   "-2147483648"
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
               Begin EditLib.fpLongInteger fpN_Num 
                  Height          =   345
                  Left            =   5220
                  TabIndex        =   21
                  Top             =   300
                  Width           =   765
                  _Version        =   196608
                  _ExtentX        =   1349
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
                  UserEntry       =   1
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
                  MaxValue        =   "9999"
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
               Begin VB.Label Label65 
                  BackStyle       =   0  '투명
                  Caption         =   "내신등급"
                  Height          =   195
                  Left            =   4410
                  TabIndex        =   233
                  Top             =   360
                  Width           =   885
               End
               Begin VB.Label Label7 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "외국어"
                  Height          =   210
                  Left            =   2430
                  TabIndex        =   232
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label62 
                  BackColor       =   &H80000001&
                  BackStyle       =   0  '투명
                  Caption         =   "탐구"
                  Height          =   210
                  Left            =   240
                  TabIndex        =   201
                  Top             =   800
                  Width           =   390
               End
               Begin VB.Label Label44 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "취소여부"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   5010
                  TabIndex        =   147
                  Top             =   45
                  Width           =   1365
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 점수"
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
                  Height          =   285
                  Left            =   60
                  TabIndex        =   98
                  Top             =   30
                  Width           =   2625
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "언어"
                  Height          =   210
                  Left            =   -390
                  TabIndex        =   97
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수리"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   930
                  TabIndex        =   96
                  Top             =   360
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '없음
            Caption         =   "Frame11"
            Height          =   3500
            Left            =   30
            TabIndex        =   86
            Top             =   90
            Width           =   8235
            Begin VB.Frame Frame3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   ">> 기본항목"
               Height          =   3450
               Left            =   30
               TabIndex        =   87
               Top             =   30
               Width           =   8175
               Begin VB.Frame Frame2 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '없음
                  Height          =   525
                  Left            =   0
                  TabIndex        =   206
                  Top             =   2430
                  Width           =   2895
                  Begin VB.OptionButton optExmN 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "무시험"
                     Height          =   285
                     Left            =   1980
                     TabIndex        =   208
                     Top             =   180
                     Width           =   885
                  End
                  Begin VB.OptionButton optExmY 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "유시험"
                     Height          =   285
                     Left            =   1080
                     TabIndex        =   207
                     Top             =   180
                     Width           =   885
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  '오른쪽 맞춤
                     BackStyle       =   0  '투명
                     Caption         =   "유/무시험"
                     Height          =   210
                     Left            =   60
                     TabIndex        =   209
                     Top             =   240
                     Width           =   975
                  End
               End
               Begin VB.OptionButton optSexMale 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "남자"
                  Height          =   285
                  Left            =   1140
                  TabIndex        =   5
                  Top             =   2100
                  Width           =   885
               End
               Begin VB.OptionButton optSexFemale 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "여자"
                  Height          =   285
                  Left            =   2010
                  TabIndex        =   6
                  Top             =   2100
                  Width           =   885
               End
               Begin VB.TextBox txt_P_Phone 
                  Height          =   270
                  Left            =   3780
                  TabIndex        =   200
                  Text            =   "Text1"
                  Top             =   2300
                  Width           =   1455
               End
               Begin VB.TextBox txt_MAJOR 
                  Height          =   285
                  Left            =   6345
                  TabIndex        =   185
                  Text            =   "txt_MAJOR"
                  Top             =   2715
                  Width           =   1725
               End
               Begin VB.TextBox txt_UNI 
                  Height          =   285
                  Left            =   3765
                  TabIndex        =   183
                  Text            =   "txt_UNI"
                  Top             =   2700
                  Width           =   1700
               End
               Begin VB.CommandButton cmdPayChg 
                  Caption         =   "결재방법"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5445
                  TabIndex        =   179
                  Top             =   3090
                  Width           =   1245
               End
               Begin VB.ComboBox cboMu_type 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   7
                  Top             =   30
                  Width           =   1725
               End
               Begin VB.ComboBox cboPTS_Sel 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   14
                  Top             =   30
                  Width           =   1275
               End
               Begin VB.CommandButton cmdNew 
                  Caption         =   "신규 (&S)"
                  Height          =   315
                  Left            =   1350
                  TabIndex        =   0
                  Top             =   -30
                  Width           =   1125
               End
               Begin VB.TextBox txtPayGbn 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   6795
                  TabIndex        =   16
                  Text            =   "txtPayGbn"
                  Top             =   3090
                  Width           =   1275
               End
               Begin VB.TextBox txtRegDate 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   3750
                  TabIndex        =   13
                  Text            =   "txtRegDate"
                  Top             =   3060
                  Width           =   1725
               End
               Begin VB.TextBox txtCel 
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   3780
                  TabIndex        =   12
                  Text            =   "txtCel"
                  Top             =   1875
                  Width           =   1455
               End
               Begin VB.TextBox txtOrdNo 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   6810
                  TabIndex        =   15
                  Text            =   "txtOrdNo"
                  Top             =   2445
                  Width           =   1275
               End
               Begin VB.TextBox txtTel 
                  Height          =   270
                  IMEMode         =   10  '한글 
                  Left            =   3780
                  TabIndex        =   11
                  Text            =   "9999-9999-9999"
                  Top             =   1560
                  Width           =   1455
               End
               Begin VB.ComboBox cboKaeyol 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   8
                  Top             =   352
                  Width           =   1725
               End
               Begin EditLib.fpMask fpExmID 
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   2
                  Top             =   750
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
               Begin VB.ComboBox cboPass4 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   56
                  Top             =   1612
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass3 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   55
                  Top             =   1192
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass2 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   54
                  Top             =   772
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass1 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   53
                  Top             =   352
                  Width           =   1275
               End
               Begin VB.ComboBox cboSel2_Sch 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   10
                  Top             =   1192
                  Width           =   1725
               End
               Begin VB.ComboBox cboSel1_Sch 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '드롭다운 목록
                  TabIndex        =   9
                  Top             =   772
                  Width           =   1725
               End
               Begin VB.TextBox txtSchNo 
                  BackColor       =   &H00C0FFFF&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   1
                  Text            =   "txtSchNo"
                  Top             =   330
                  Width           =   1605
               End
               Begin VB.TextBox txtStdNM 
                  Height          =   345
                  IMEMode         =   10  '한글 
                  Left            =   1140
                  TabIndex        =   3
                  Text            =   "txtStdNM"
                  Top             =   1170
                  Width           =   1605
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '없음
                  Height          =   435
                  Left            =   1140
                  TabIndex        =   88
                  Top             =   2025
                  Width           =   1800
               End
               Begin EditLib.fpMask fpBirth_ymd 
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   4
                  Top             =   1590
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
                  Mask            =   "9999-99-99"
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
               Begin VB.Label Label63 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "성별"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   205
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label61 
                  Caption         =   "학부모HP"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   199
                  Top             =   2280
                  Width           =   880
               End
               Begin VB.Label Label60 
                  BackStyle       =   0  '투명
                  Caption         =   "지원단대"
                  Height          =   300
                  Left            =   5520
                  TabIndex        =   184
                  Top             =   2760
                  Width           =   750
               End
               Begin VB.Label Label59 
                  Caption         =   "지원대학"
                  Height          =   225
                  Left            =   2970
                  TabIndex        =   182
                  Top             =   2760
                  Width           =   810
               End
               Begin VB.Label Label53 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "등급"
                  Height          =   210
                  Left            =   2520
                  TabIndex        =   159
                  Top             =   105
                  Width           =   1185
               End
               Begin VB.Label Label52 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수리점수구분"
                  Height          =   210
                  Left            =   5550
                  TabIndex        =   158
                  Top             =   105
                  Width           =   1185
               End
               Begin VB.Label Label42 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "등록일자"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   2310
                  TabIndex        =   146
                  Top             =   3105
                  Width           =   1365
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "핸드폰"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2730
                  TabIndex        =   144
                  Top             =   1890
                  Width           =   975
               End
               Begin VB.Label Label40 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "주문번호(조회)"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   5370
                  TabIndex        =   143
                  Top             =   2490
                  Width           =   1365
               End
               Begin VB.Label Label39 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "TEL"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2730
                  TabIndex        =   142
                  Top             =   1620
                  Width           =   975
               End
               Begin VB.Label Label28 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "계  열"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   126
                  Top             =   390
                  Width           =   975
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "4지망 합격학원"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   117
                  Top             =   1650
                  Width           =   1455
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "3지망 합격학원"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   116
                  Top             =   1230
                  Width           =   1455
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "2지망 합격학원"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   115
                  Top             =   810
                  Width           =   1455
               End
               Begin VB.Label Label18 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "1지망 합격학원"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   114
                  Top             =   390
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "2지망 학원"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   113
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label16 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "1지망 학원"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   112
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  '투명
                  Caption         =   ">> 기본항목"
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
                  Height          =   285
                  Left            =   90
                  TabIndex        =   93
                  Top             =   60
                  Width           =   2625
               End
               Begin VB.Label Label4 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "시스템 코드"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   92
                  Top             =   390
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수험번호"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   91
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "학생명"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   90
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "생년월일"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   89
                  Top             =   1650
                  Width           =   975
               End
            End
         End
         Begin VB.Label Label45 
            BackStyle       =   0  '투명
            Caption         =   "※ 학생삭제는 잘못 입력한 경우만 사용하십시요."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3810
            TabIndex        =   148
            Top             =   10770
            Width           =   4365
         End
      End
   End
   Begin FPSpread.vaSpread sprStdData 
      Height          =   165
      Left            =   8430
      TabIndex        =   145
      Top             =   9300
      Width           =   2595
      _Version        =   393216
      _ExtentX        =   4577
      _ExtentY        =   291
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
      SpreadDesigner  =   "STD010_N.frx":8E75
   End
   Begin VB.Label Label64 
      BackColor       =   &H80000001&
      BackStyle       =   0  '투명
      Caption         =   "탐구"
      Height          =   210
      Left            =   3780
      TabIndex        =   231
      Top             =   4650
      Width           =   390
   End
End
Attribute VB_Name = "STD010_N"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD010_N
'   모 듈  목 적 : 학생 등록 및 조회
'
'   작   성   일 : 2007/08/22
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

'테스트 바꿔봐요
Private Type tExcel_StdData
    ACID        As String
    EXMID       As String
    STDNM       As String
    Birth_ymd       As String
    EXMTYPE     As String
    kaeyol      As String
    WANT_ACID1  As String
    WANT_ACID2  As String
    KOR         As Long
    ENG         As Long
    MAT         As Long
    
    SATAM1      As String
    SATAM2      As String
    SATAM3      As String
    SATAM4      As String
    SATAM5      As String
    SATAM6      As String
    SATAM7      As String
    SATAM8      As String
    SATAM9      As String
    SATAM10     As String
    
    ENG2        As String
    
    GWATAM1     As String
    GWATAM2     As String
    GWATAM3     As String
    GWATAM4     As String
    GWATAM5     As String
    GWATAM6     As String
    GWATAM7     As String
    GWATAM8     As String
    
    SURI        As String
    
    NONSUL1     As String
    NONSUL2     As String
    NONSUL3     As String
    NONSUL4     As String
End Type
Private uExcel_StdData      As tExcel_StdData



Private Sub cbo_Tamgoo_Click()
    If cbo_Tamgoo.ListIndex = 0 Then
        fpT_Num = ""
    End If
End Sub


Private Sub cmdClinicClear_Click()
    ' 클리닉들 초기화
    Call Init_Clinic(chkClinic_L, chkClinic_M, chkClinic_E)
End Sub

' 클리닉창 확인
Private Sub cmdClinicOK_Click()
    fraClinic.Visible = False
End Sub

'클리닉 창 보여주기
Private Sub cmdSelClinic_Click()
    fraClinic.Visible = True
    
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub



Private Sub Form_Load()
        
    Me.Move 0, 0, 15255, 11715
    
    fraGwamok.Visible = False '폼 아래 과목코드 프레임


    '>>>>>>>>>>>>등록 폼 초기화 <<<<<<<<<<<<<<
    txtSchNo.Text = ""
    fpExmID.Text = ""
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""
    txt_P_Phone = ""
    txt_UNI.Text = ""
    txt_MAJOR.Text = ""
    txtTel.Text = ""
    txtCel.Text = ""
    txtRegDate.Text = ""
    txtOrdNo.Text = ""
    txtPayGbn.Text = ""

    txtCancel.Text = ""
    cbo_Tamgoo.ListIndex = -1

    optExmY.value = True
    optSexMale.value = True
    
    fpK_Num.value = 0
    fpE_Num.value = 0
    fpM_Num.value = 0
    fpN_Num.value = 0

    '작은창.
    fpBirth_ymdS.Text = ""
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtEmail.Text = ""
    
    Call basCommonSTD.Init_CboKaeyolDefault(cboKaeyol)      '계열
    Call basCommonSTD.Init_CboSch(cboSel1_Sch)   '1지망 학원
    Call basCommonSTD.Init_CboSch(cboSel2_Sch)   '2지망 학원
    Call basCommonSTD.Init_CboSch(cboPass1)      '1지망 합격 학원
    Call basCommonSTD.Init_CboSch(cboPass2)      '2지망 합격 학원
    Call basCommonSTD.Init_CboSch(cboPass3)      '3지망 합격 학원
    Call basCommonSTD.Init_CboSch(cboPass4)      '4지망 합격 학원
    Call basCommonSTD.Init_Mu_type(cboMu_type)       '등급
    Call basCommonSTD.Init_PTS_Sel(cboPTS_Sel)       '수리점수구분
    Call basCommonSTD.Init_Card(cboCard, Get_CodeCardBySchool(basModule.schcd))             '카드
    
    
    '>>>>>>>>>>>> 조회 폼 초기화 <<<<<<<<<<<<<<
    fpExmID_F.Text = ""
    fpExmID_E.Text = ""
    
    txtStdNM_F.Text = ""
    fpBirth_ymd_F.Text = ""
    sprSTD_F.MaxRows = 0
    
    sprExcel_STD_Data.MaxRows = 0
    
    fpPayOK.value = 0
    fpPayNot.value = 0
    fpPayTot.value = 0
    
    
    Call basCommonSTD.Init_CboKaeyolDefault(cboKaeyol_F)    '조회 계열
    cboKaeyol_F.AddItem "전체" & Space(30) & "ALL", 0
    cboKaeyol_F.ListIndex = 0
    
    Call basCommonSTD.Init_InGbn(cboinGbn)           '조회 인터넷/학원
    Call basCommonSTD.Init_ExmType(cboExmType)       '조회 유무험시험
    Call basCommonSTD.Init_Pay(cboPay)               '조회 결제
    Call basCommonSTD.Init_PassCN(cboPassCN)         '조회 합격차수
    Call basCommonSTD.Init_CboSch(cboSel1_SCH_F)        '조회 1지망 학원
    Call basCommonSTD.Init_CboSch(cboSel2_SCH_F)        '조회 2지망 학원
    
    Call basCommonSTD.Set_Spread_Design1(sprSTD_F)              '학생조회 시트
    Call basCommonSTD.Set_Spread_Design1(sprExcel_STD_Data)     '엑셀가져오기 시트
    
    ' 노량진,송파 폼의 탐구 콤보박스
    With cbo_Tamgoo
        .AddItem "없음                    X"
        .AddItem constSatams(0) & "                    " & constSatamCodes(0)
        .AddItem constSatams(1) & "                    " & constSatamCodes(1)
        .AddItem constSatams(2) & "                    " & constSatamCodes(2)
        .AddItem constSatams(3) & "                    " & constSatamCodes(3)
        .AddItem constSatams(4) & "                    " & constSatamCodes(4)
        .AddItem constSatams(5) & "                    " & constSatamCodes(5)
        .AddItem constSatams(6) & "                    " & constSatamCodes(6)
        .AddItem constSatams(7) & "                    " & constSatamCodes(7)
        .AddItem constSatams(8) & "                    " & constSatamCodes(8)
        .AddItem constSatams(9) & "                    " & constSatamCodes(9)
        .AddItem "세지                    11"
        .AddItem "물I                    51"
        .AddItem "화I                    52"
        .AddItem "생I                    53"
        .AddItem "지I                    54"
        .AddItem "물II                    55"
        .AddItem "화II                    56"
        .AddItem "생II                    57"
        .AddItem "지II                    58"
    End With
    
    '>>>>>>>>>>>> 작은창들 위치세팅. <<<<<<<<<<<<<<
    With fraAddr        '< 학생 상세내역 등록 : 2008.01.10
        .Top = 3420
        .Left = 6540
        
        .ZOrder 0
        .Visible = False
    End With
    With fraPoint       '< 학생 상세점수 등록 : 2008.01.10
        .Top = 3420
        .Left = 4500

        .ZOrder 0
        .Visible = False
    End With
    
    With FraPay         '< 결재정보 등록 : 2010.01.13
'        .Top = 3420
'        .Left = 4500
    
        .ZOrder 0
        .Visible = False
    End With
    
    ' 클리닉
    With fraClinic
        .Top = 9295.282
        .Left = 3584.906
        
        .ZOrder 0
        .Visible = False
    End With
    
    
    '>>>>>>>>>>>> 학원에 따른 폼 설정 <<<<<<<<<<<<<<
    Dim ni As Integer
    
    '>> 1지망 학원
    Call basCommonSTD.Set_CboSch(cboSel1_Sch, basModule.schcd)
        
    '>> 학원
    Select Case Trim(basModule.schcd)
        Case "N"        '노량진
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            chkEng2(13).Visible = True
            
        Case "M"        '강남 마이맥
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            chkEng2(13).Visible = True
            
        Case "S"        '송파
            
            chkSatam(1).Visible = True
            chkSatam(2).Visible = True
            chkSatam(3).Visible = True
            chkSatam(4).Visible = True
            
            'chkSatam(5).Visible = False
            
            chkSatam(5).Visible = True
            chkSatam(6).Visible = True
            chkSatam(7).Visible = True
            chkSatam(8).Visible = True
            chkSatam(9).Visible = True
            'chkSatam(11).Visible = False
            
            chkEng2(3).Visible = False
            
            chkEng2(7).Visible = True
            chkEng2(8).Visible = True
            
            chkEng2(8).Caption = "논술"     ''' 이부분은 고쳐야한다 송파....
            chkNonsul(1).Visible = False
            chkNonsul(2).Caption = "언수외"
            chkNonsul(3).Caption = "논술"
            
            chkEng2(9).Visible = True
            
            chkEng2(10).Visible = False     ' True
            chkEng2(11).Visible = False     ' True
            
            chkEng2(12).Visible = True
            chkEng2(13).Visible = True
            
        Case "J"        '양재
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            chkEng2(13).Visible = True
            
        Case Else
            For ni = 7 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            chkEng2(13).Visible = True
    End Select
    
    
    If basModule.schcd = "N" Or basModule.schcd = "S" Then
        cmdSelClinic.Visible = True
    End If
    
    If g_sClinic_Ls(0) = "" Then
        Call basGwamok.setConstant '다시 초기화
    End If
    
    ' 클리닉창 세팅
    For ni = 0 To CLINIC_L_COUNT - 1
        chkClinic_L(ni).Caption = g_sClinic_Ls(ni)
        chkClinic_L(ni).Visible = True
    Next ni
    For ni = 0 To CLINIC_M_COUNT - 1
        chkClinic_M(ni).Caption = g_sClinic_Ms(ni)
        chkClinic_M(ni).Visible = True
    Next ni
    For ni = 0 To CLINIC_E_COUNT - 1
        chkClinic_E(ni).Caption = g_sClinic_Es(ni)
        chkClinic_E(ni).Visible = True
    Next ni
    
    
    
    
End Sub


'>> 신규
Private Sub cmdNew_Click()

    Dim ni      As Integer

    txtSchNo.Text = ""
    fpExmID.Text = ""
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""
    txt_P_Phone = ""
    txt_UNI.Text = ""
    txt_MAJOR.Text = ""
    txtTel.Text = ""
    txtCel.Text = ""
    txtRegDate.Text = ""
    
    cbo_Tamgoo.ListIndex = -1
    cboPass1.ListIndex = 0
    cboPass2.ListIndex = 0
    cboPass3.ListIndex = 0
    cboPass4.ListIndex = 0

    txtOrdNo.Text = ""
    txtPayGbn.Text = ""
    txtCancel.Text = ""
    fpT_Num.Text = ""   '탐구점수

    optExmY.value = True
    optSexMale.value = True
    
    fpK_Num.value = 0
    fpE_Num.value = 0
    fpM_Num.value = 0
    fpN_Num.value = 0

    '>>>>>작은창.
    fpBirth_ymdS.Text = ""
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtEmail.Text = ""
    
    '>> 클리닉 창 초기화
    Call Init_Clinic(chkClinic_L, chkClinic_M, chkClinic_E)
    
    cboMu_type.ListIndex = cboMu_type.ListCount - 1
    cboKaeyol.ListIndex = 0
    
    
    For ni = 1 To SATAM_COUNT + 1 Step 1
        chkSatam(ni).value = 0
    Next ni

    For ni = 1 To ENG2_COUNT Step 1
        chkEng2(ni).value = 0
    Next ni

    For ni = 1 To 8 Step 1
        chkGwatam(ni).value = 0
    Next ni

    For ni = 1 To 4 Step 1
        chkMath(ni).value = 0
        chkNonsul(ni).value = 0
    Next ni


    '>> 1지망 학원
    Call basCommonSTD.Set_CboSch(cboSel1_Sch, basModule.schcd)

    '>> 2지명 학원
    cboSel2_Sch.ListIndex = 0
    
End Sub


'선택 및 조회 된 내용에 대한 화면 변경
Private Sub changeEnableGwamoks(bSatam As Boolean, bEng2 As Boolean, bGwatam As Boolean, bMath As Boolean, bNonsul As Boolean)

    Dim ni      As Integer
    
    
    For ni = 1 To SATAM_COUNT + 1 Step 1    '< 사탐
        If True = bSatam Then
            chkSatam(ni).Enabled = True
        Else
            chkSatam(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To ENG2_COUNT Step 1                 '< 제2외국어
        If True = bEng2 Then
            chkEng2(ni).Enabled = True
        Else
            chkEng2(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 8 Step 1                  '< 과탐
        If True = bGwatam Then
            chkGwatam(ni).Enabled = True
        Else
            chkGwatam(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 4 Step 1                  '< 수리
        If True = bMath Then
            chkMath(ni).Enabled = True
        Else
            chkMath(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 4 Step 1                  '< 논술
        If True = bNonsul Then
            chkNonsul(ni).Enabled = True
        Else
            chkNonsul(ni).Enabled = False
        End If
    Next ni
    
    
End Sub


Private Sub cboKaeyol_Click()

    If Me.Tag = "LOAD" Then Exit Sub
    
    ' 과목 체크박스 상태값 변경
    Select Case Trim(basModule.schcd)
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "07", "09", "11", "13", "21"
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "02", "04", "08", "10", "12", "14", "22"
                    Call changeEnableGwamoks(False, False, True, True, True)
                    
                Case "05", "15"
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "06", "16"
                    Call changeEnableGwamoks(False, False, True, True, True)
            End Select
        Case "S", "P", "J"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "05", "11", "18"                                                         '< 2008.02.15 : 계열 - 송파, 마송, 양재      2009.06.02 : 계열추가
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "02", "04", "06", "08", "12", "19"                                                   '< 2008.02.15 : 계열 - 송파, 마송, 양재      2009.06.02 : 계열추가
                    Call changeEnableGwamoks(False, False, True, True, True)
            End Select
        Case Else
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "04", "06", "11", "16"                               '< 2008.01.10 : 계열 - 강남
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "02", "05", "07", "12", "17"                                     '< 2008.01.10 : 계열 - 강남
                    Call changeEnableGwamoks(False, False, True, True, True)
            End Select
    End Select
    
    
    
End Sub




'Private Sub cboKaeyol_Click()
'    If Me.Tag = "LOAD" Then Exit Sub
'    changeKaeyol
'End Sub




Private Sub cmdGwamokView_Click()
    fraGwamok.Left = 60
    fraGwamok.Top = 3390
    fraGwamok.ZOrder 0
    
    fraGwamok.Visible = True
End Sub

Private Sub cmdClose_Click()
    fraGwamok.Visible = False
End Sub




'>> 학생등록하기
Private Sub cmdStdin_Click()
    Dim bRet        As Boolean
    
    '>> 체크조건
    If Trim(fpExmID.UnFmtText) = "" Then
        MsgBox "수험번호가 없습니다.", vbExclamation + vbOKOnly, "학생등록"
        Exit Sub
    End If
    If Trim(fpBirth_ymd.UnFmtText) = "" Then
        MsgBox "생년월일이 없습니다.", vbExclamation + vbOKOnly, "학생등록"
        Exit Sub
    End If
    
    
    On Error GoTo ErrStmt
    
    cmdStdin.Enabled = False
        bRet = Save_Stdin           '<< 학생등록
            
    cmdStdin.Enabled = True
    
    If bRet = True Then
        MsgBox "학생 등록하였습니다.", vbInformation + vbOKOnly, "학생등록"
        Call cmdNew_Click
        
    Else
        MsgBox "학생 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생등록"
    End If
    
    Exit Sub
ErrStmt:

    MsgBox "학생등록시 오류가 발생하였습니다." & vbCrLf & _
        Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "학생등록"
        
    On Error GoTo 0
    
End Sub

'>> 학생등록
Private Function Save_Stdin() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    Dim nExe        As Integer
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    
    '>> 학생 등록/갱신
        sTmp = "INSERT":    If Trim(txtSchNo.Text) > "" Then sTmp = "UPDATE"
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_STYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> 시스템코드
        sTmp = Trim(txtSchNo.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 학원코드
        sTmp = basModule.schcd
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 수험번호
        sTmp = Trim(fpExmID.UnFmtText)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 학생명
        sTmp = Trim(txtStdNM.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 생년월일
        sTmp = Trim(fpBirth_ymd.UnFmtText)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("v_birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> 유/무시험 구분
        If optExmY.value = True Then
            sTmp = "1"
        ElseIf optExmN.value = True Then
            sTmp = "0"
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_EXMTYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> 계열
        sTmp = Trim(Right(cboKaeyol.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_KAEYOL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    
    '## 선택과목 ###
        '>> 사탐과목 선택
        sTmp = ""
        For ni = 1 To SATAM_COUNT Step 1
            If chkSatam(ni).value = 1 Then
                sTmp = sTmp & Format(SATAM_CLASS + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 제2외국어 선택
        sTmp = ""
        For ni = 1 To ENG2_COUNT Step 1         '< 2008.01.14 : 송파 추가내역
            If chkEng2(ni).value = 1 Then
                If ni = 13 Then '베트남어 코드구함 :   베트남어코드:44 수리나형:43 이라서 +1해줌
                    sTmp = sTmp & Format(31 + ni, "00") & "|"
                Else
                    sTmp = sTmp & Format(30 + ni, "00") & "|"
                End If
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 과탐과목 선택
        sTmp = ""
        For ni = 1 To 8 Step 1
            If chkGwatam(ni).value = 1 Then
                sTmp = sTmp & Format(50 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 수리과목 선택
        sTmp = ""
        For ni = 1 To 4 Step 1
            If chkMath(ni).value = 1 Then
                sTmp = sTmp & Format(80 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 논술과목 선택
        sTmp = ""
        For ni = 1 To 4 Step 1
            If chkNonsul(ni).value = 1 Then
                sTmp = sTmp & Format(90 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL5", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
        '>> 클리닉 선택
        '국어, 수학, 영어 레디오박스
        sTmp = ""
        sTmp = sTmp & IIf(Get_IndexByChk(chkClinic_L) = -1, "", CStr(CLINIC_L_CLASS + Get_IndexByChk(chkClinic_L)) & "|")
        sTmp = sTmp & IIf(Get_IndexByChk(chkClinic_M) = -1, "", CStr(CLINIC_M_CLASS + Get_IndexByChk(chkClinic_M)) & "|")
        sTmp = sTmp & IIf(Get_IndexByChk(chkClinic_E) = -1, "", CStr(CLINIC_E_CLASS + Get_IndexByChk(chkClinic_E)) & "|")
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL7", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        

    '>> 국어점수
        nTmp = CDbl(fpK_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_K_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> 영어점수
        nTmp = CDbl(fpE_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_E_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> 수학점수
        nTmp = CDbl(fpM_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_M_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> 합계
        nTmp = CDbl(fpK_Num.value) + CDbl(fpM_Num.value) + CDbl(fpE_Num.value) + CDbl(fpT_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> 내신등급
        nTmp = CDbl(fpN_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_N_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    '>> 1지망 학원
        sTmp = Trim(Right(cboSel1_Sch.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL1_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 2지망 학원
        sTmp = ""
        If Trim(Right(cboSel2_Sch.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboSel2_Sch.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL2_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


    '>> 1지망 합격학원
        sTmp = ""
        If Trim(Right(cboPass1.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass1.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 2지망 합격학원
        sTmp = ""
        If Trim(Right(cboPass2.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass2.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 3지망 합격학원
        sTmp = ""
        If Trim(Right(cboPass3.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass3.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 4지망 합격학원
        sTmp = ""
        If Trim(Right(cboPass4.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass4.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> 전화번호
        sTmp = Trim(txtTel.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 핸드폰
        sTmp = Trim(txtCel.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_CEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> 수리점수구분
        sTmp = ""
        If Trim(Right(cboPTS_Sel.Text, 30)) <> "X" Then sTmp = Trim(Right(cboPTS_Sel.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PTS", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 등급
        sTmp = ""
        If Trim(Right(cboMu_type.Text, 30)) <> "X" Then sTmp = Trim(Right(cboMu_type.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PTS", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> 학부모HP
        sTmp = Trim(txt_P_Phone.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PRNT_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> 지원대학
        sTmp = Trim(txt_UNI.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_D_UNIVCD", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> 지원단대
        sTmp = Trim(txt_MAJOR.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_D_MAJORCD", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
    '>> 탐구
        sTmp = ""
        If Trim(Right(cbo_Tamgoo.Text, 30)) <> "X" Then sTmp = Trim(Right(cbo_Tamgoo.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL6", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> 탐구점수
        nTmp = CDbl(fpT_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_T_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
    '>> 성별 구분
        If optSexMale.value = True Then
            sTmp = SexMaleValue
        ElseIf optSexFemale.value = True Then
            sTmp = SexFemaleValue
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEX", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        
    '>> 데이터 등록
    DBCmd.CommandType = adCmdStoredProc
    DBCmd.CommandText = "PG_STD.PROC_STD_SAVE"
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    Save_Stdin = True
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    
    
    
    
    Exit Function
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_Stdin = False
    
End Function


'>> 학생삭제하기
Private Sub cmdStdDel_Click()
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    '>> 체크조건
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "시스템코드가 없습니다.", vbExclamation + vbOKOnly, "학생삭제"
        Exit Sub
    End If
    
    sTmp = Trim(txtStdNM.Text) & "의 학생을 삭제하시겠습니까?"
    If MsgBox(sTmp, vbQuestion + vbYesNo, "학생삭제") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdStdDel.Enabled = False
        bRet = Delete_StdOut
        
    cmdStdDel.Enabled = True
    
    If bRet = True Then
        MsgBox "학생 삭제하였습니다.", vbInformation + vbOKOnly, "학생삭제"
    Else
        MsgBox "학생 삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생삭제"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "학생삭제시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생삭제"
    On Error GoTo 0
    
End Sub

'>> 학생삭제
Private Function Delete_StdOut() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
            
    '>> 시스템코드
        sTmp = Trim(txtSchNo.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 학원코드
        sTmp = basModule.schcd
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> 데이터 등록
    DBCmd.CommandType = adCmdStoredProc
    DBCmd.CommandText = "PG_STD.PROC_STD_DELETE"
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    Delete_StdOut = True
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_StdOut = False
End Function






Private Sub cmdCancel_Click()
    
    Dim bRet        As Boolean
    
    '>> 체크조건
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "시스템 코드가 없습니다.", vbExclamation + vbOKOnly, "학생 취소하기"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdCancel.Enabled = False
        bRet = Cancel_StdOut
        
    cmdCancel.Enabled = True
    
    If bRet = True Then
        MsgBox "학생 합격취소 하였습니다.", vbInformation + vbOKOnly, "학생 취소하기"
    Else
        MsgBox "학생 합격취소시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생 취소하기"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "학생 합격취소시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생 취소하기"
    On Error GoTo 0
    
End Sub

'>> 합격취소하기
Private Function Cancel_StdOut() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    Dim sStr        As String
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nExe        As Integer
    
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
  
  
    nExe = 0
    
    sStr = ""
    sStr = sStr & " INSERT INTO CLSTD91TB"
    sStr = sStr & " SELECT *"
    sStr = sStr & "   FROM CLSTD01TB "
    sStr = sStr & "   WHERE SCHNO   = '" & Trim(txtSchNo.Text) & "'"
    
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
        nExe = 0
        
        '-----------------------------------------------------------------------------------------------------
        Select Case Trim(basModule.schcd)
            Case "S"
                sStr = ""
                sStr = sStr & " UPDATE CLSTD01TB "
                sStr = sStr & "    SET EXMID   = '',"
                sStr = sStr & "        PASS1   = '',"
                sStr = sStr & "        PASS2   = '',"
                sStr = sStr & "        PASS3   = '',"
                sStr = sStr & "        PASS4   = '',"
                sStr = sStr & "        CY_ACNT = '',"
                sStr = sStr & "        TOT_AMT = 0 "
                sStr = sStr & "  WHERE SCHNO   = '" & Trim(txtSchNo.Text) & "'"
            Case Else
                sStr = ""
                sStr = sStr & " DELETE "
                sStr = sStr & "   FROM CLSTD01TB "
                sStr = sStr & "  WHERE SCHNO   = '" & Trim(txtSchNo.Text) & "'"
        End Select
        
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
        
        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
        
        If nExe = 1 Then
            nExe = 0
            On Error Resume Next
            
            sStr = ""
            sStr = sStr & " INSERT INTO CLSTD92TB (SCHNO, ACID, EXMID, TIMESTAMP) "
            sStr = sStr & " VALUES( "
            sStr = sStr & "         '" & Trim(txtSchNo.Text) & "',"
            sStr = sStr & "         '" & Trim(basModule.schcd) & "',"
            sStr = sStr & "         '" & Trim(fpExmID.UnFmtText) & "',"
            sStr = sStr & "         SYSDATE"
            sStr = sStr & "       ) "
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
            
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            If nExe = 1 Then
                bRet = True
            Else
                nExe = 0
                
                On Error GoTo 0
                On Error GoTo ErrStmt
                
                sStr = ""
                sStr = sStr & " UPDATE CLSTD92TB "
                sStr = sStr & "    SET ACID  = '" & Trim(basModule.schcd) & "',"
                sStr = sStr & "        EXMID = '" & Trim(fpExmID.UnFmtText) & "',"
                sStr = sStr & "        TIMESTAMP = SYSDATE "
                sStr = sStr & "  WHERE SCHNO = '" & Trim(txtSchNo.Text) & "'"
                
                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
                
                DBCmd.Execute nExe, , -1
                
                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop
                
                If nExe = 1 Then
                    bRet = True
                End If
                
            End If
        End If
        '-----------------------------------------------------------------------------------------------------
        
    End If
        
    
    Cancel_StdOut = bRet
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Cancel_StdOut = bRet
End Function
















'>> 학생 조회하기
Private Sub cmdFind_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim sGbn        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    cmdFind.Enabled = False
    
    sprSTD_F.MaxRows = 0
    fpPayOK.value = 0
    fpPayNot.value = 0
    fpPayTot.value = 0
    
    '2011-01-11 김한욱 노량진 요청에 의해 엑셀 제일 뒷 부분에 지원대학 및 지원 단대 입력
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, EXMID, STDNM, SEL1_SCH , SEL2_SCH, "
    
    sStr = sStr & " birth_ymd, "
    
    '<< 계열 >> : 2012.12.10
    sStr = sStr & basCommonSTD.Get_SqlKaeyolDecode
    
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* 사탐-한국사 */"
    sStr = sStr & "             '" & constSatamCodes(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "             '" & constSatamCodes(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* 사탐-동아시아사 */"
    sStr = sStr & "             '" & constSatamCodes(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생명과학1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "             '" & constSatamCodes(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* 사탐-세계지리 */"
    sStr = sStr & "             '" & constSatamCodes(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* 사탐-생활과윤리 */"
    sStr = sStr & "             '" & constSatamCodes(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* 사탐-윤리와사상 */"
    sStr = sStr & "             '" & constSatamCodes(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생명과학2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* 사탐-법과정치 */"
    sStr = sStr & "             '" & constSatamCodes(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "             '" & constSatamCodes(8) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* 사탐-사회문화 */"
    sStr = sStr & "             '" & constSatamCodes(9) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL10,"
    sStr = sStr & "  "
    sStr = sStr & "      /* 제2외국어 & 수리 */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '31'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '32'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '33'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '34'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '35'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '36'"
    '<< 송파 >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '37'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '38'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '39'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '40'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '41'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '81'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '82'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '83'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '84'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END END END END END END END END END END END END END END SEL_X2,"
    sStr = sStr & "  "
    sStr = sStr & "      /* 논술 */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
    sStr = sStr & "             '91'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N1,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
    sStr = sStr & "             '92'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N2,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 사탐 */"
    sStr = sStr & "             '93'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /* 과탐 */"
    sStr = sStr & "             '94'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         PAYOK, PAYNOT, "
    sStr = sStr & "         GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.schcd) & "') AS PAYTOT, "        '< 전체집계 하는 함수
    sStr = sStr & "      /* 탐구 성적 떄문에 처리.. */"
    sStr = sStr & "              CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(0) & "') > 0 THEN '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(1) & "') > 0 THEN '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(2) & "') > 0 THEN '" & constSatams(2) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(3) & "') > 0 THEN '" & constSatams(3) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(4) & "') > 0 THEN '" & constSatams(4) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(5) & "') > 0 THEN '" & constSatams(5) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(6) & "') > 0 THEN '" & constSatams(6) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(7) & "') > 0 THEN '" & constSatams(7) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(8) & "') > 0 THEN '" & constSatams(8) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(9) & "') > 0 THEN '" & constSatams(9) & " '"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'51') > 0 THEN '물I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'52') > 0 THEN '화I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'53') > 0 THEN '생I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'54') > 0 THEN '지I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'55') > 0 THEN '물II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'56') > 0 THEN '화II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'57') > 0 THEN '생II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'58') > 0 THEN '지II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 IS NULL THEN ''"
    sStr = sStr & "         END END END END END END END END END END END END END END END END END END END SEL_X6,"
    sStr = sStr & "         K_NUM, M_NUM, E_NUM, N_NUM, T_NUM, TOT_NUM, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2, TEL, CEL, "
    sStr = sStr & "         REGDATE, PAYGBN, CASH_BILL_NUM, D_UNIVCD, D_MAJORCD, PRNT_TEL "
    
    Select Case Trim(Right(cboPassCN, 30))
        Case "ALL"      ' /* 합격생은 모두 조회 */
            sStr = sStr & " FROM (SELECT SCHNO, MAX(EXMID) AS EXMID, MAX(STDNM) AS STDNM, MAX(SEL1_SCH) AS SEL1_SCH, MAX(SEL2_SCH) AS SEL2_SCH, MAX(D_UNIVCD) AS D_UNIVCD, MAX(D_MAJORCD) AS D_MAJORCD, MAX(Birth_ymd) AS Birth_ymd,"
            sStr = sStr & "              MAX(KAEYOL) AS KAEYOL,"
            sStr = sStr & "              MAX(SEL1) AS SEL1, MAX(SEL2) AS SEL2, MAX(SEL3) AS SEL3, MAX(SEL4) AS SEL4, MAX(SEL5) SEL5, MAX(SEL6) SEL6, "
            sStr = sStr & "              MAX(CL_CLOSE) AS CL_CLOSE, "
            sStr = sStr & "              MAX(PAYOK) AS PAYOK, MAX(PAYNOT) AS PAYNOT, "
            sStr = sStr & "              MAX(K_NUM) AS K_NUM, MAX(M_NUM) AS M_NUM, MAX(E_NUM) AS E_NUM, MAX(N_NUM) AS N_NUM, MAX(T_NUM) AS T_NUM, MAX(TOT_NUM) AS TOT_NUM, "
            sStr = sStr & "              MAX(ZIP) AS ZIP, MAX(ADDR1) AS ADDR1, MAX(ADDR2) AS ADDR2, MAX(TEL) AS TEL, MAX(CEL) AS CEL, "
            sStr = sStr & "              MAX(REGDATE) AS REGDATE, MAX(PAYGBN) AS PAYGBN, MAX(CASH_BILL_NUM) AS CASH_BILL_NUM, MAX(PRNT_TEL) AS PRNT_TEL "
            sStr = sStr & "         FROM ("
            '==========================================================================================================
            
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, Birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, SEL6, CL_CLOSE, "
            sStr = sStr & "                      PAYOK, PAYNOT, "
            sStr = sStr & "                      NVL(K_NUM,0) AS K_NUM, NVL(M_NUM,0) AS M_NUM, NVL(E_NUM,0) AS E_NUM, NVL(N_NUM,0) AS N_NUM, NVL(T_NUM,0) AS T_NUM,"
            'sStr = sStr & "                      (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS TOT_NUM ,"
            sStr = sStr & "                      TOT_NUM ,"
            sStr = sStr & "                      SUBSTR(A.ZIP,1,3)||'-'||SUBSTR(A.ZIP,4,3) AS ZIP, A.ADDR1, A.ADDR2, A.TEL, A.CEL, "
            sStr = sStr & "                      TO_CHAR(A.REGDATE,'YYYY-MM-DD HH24:MI') AS REGDATE, GET_PAYGUBN(A.ORD_NO) AS PAYGBN, CASH_BILL_NUM, PRNT_TEL AS PRNT_TEL "
            sStr = sStr & "                 From CLSTD01TB A, "
            sStr = sStr & "                      ("
            sStr = sStr & "                       SELECT ACID, SUM(PAYOK) AS PAYOK, SUM(PAYNOT) AS PAYNOT"
            sStr = sStr & "                         FROM ("
            sStr = sStr & "                               SELECT ACID, "
            sStr = sStr & "                                      CASE WHEN EXMID > ' ' THEN"
            sStr = sStr & "                                          1"
            sStr = sStr & "                                      Else"
            sStr = sStr & "                                          0"
            sStr = sStr & "                                      END PAYOK,"
            sStr = sStr & "                                      CASE WHEN EXMID IS NULL THEN"
            sStr = sStr & "                                          1"
            sStr = sStr & "                                      Else"
            sStr = sStr & "                                          0"
            sStr = sStr & "                                      END PAYNOT"
            sStr = sStr & "                                 FROM CLSTD01TB "
            
            sStr = sStr & "                                WHERE ACID = '" & Trim(basModule.schcd) & "'"
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "                              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "                              AND EXMTYPE = '1'"
            End If
            
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "                              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "                              AND R_WAY IN ('1','3') "
            End If
            
            '<< 결재여부 >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "                          AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "                          AND EXMID IS NULL "
            End Select
            
            If Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) <> "" Then
                sStr = sStr & "                              AND EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
                sStr = sStr & "                                            AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            ElseIf Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) = "" Then
                sStr = sStr & "                              AND EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
                sStr = sStr & "                                            AND '99999'"
            ElseIf Trim(fpExmID_F.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) <> "" Then
                sStr = sStr & "                              AND EXMID BETWEEN '00000'"
                sStr = sStr & "                                            AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            Else
                'no action
            End If
            
            If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then      ' 인문
                sStr = sStr & "                              AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
            End If
    
            If Trim(txtStdNM_F.Text) <> "" Then
                sStr = sStr & "                              AND STDNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
            End If
            If Trim(fpBirth_ymd_F.UnFmtText) <> "" Then
                sStr = sStr & "                              AND Birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
            End If
            If Trim(Right(cboSel1_SCH_F.Text, 30)) <> "X" Then
                sStr = sStr & "                              AND SEL1_SCH = '" & Trim(Right(cboSel1_SCH_F.Text, 30)) & "'"
            End If
            If Trim(Right(cboSel2_SCH_F.Text, 30)) <> "X" Then
                sStr = sStr & "                              AND SEL2_SCH = '" & Trim(Right(cboSel2_SCH_F.Text, 30)) & "'"
            End If
            
            sStr = sStr & "                                  AND CL_CLOSE IS NULL "
            sStr = sStr & "                                  AND BIGO2 IS NULL "
            
            sStr = sStr & "                              )"
            sStr = sStr & "                         GROUP BY ACID"
            sStr = sStr & "                      ) B"
            sStr = sStr & "                WHERE A.ACID = B.ACID"
            sStr = sStr & "                  AND A.ACID = '" & Trim(basModule.schcd) & "'"
            
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "              AND EXMTYPE = '1'"
            End If
            
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "              AND R_WAY IN ('1','3') "
            End If
            '<< 결재여부 >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "          AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "          AND EXMID IS NULL "
            End Select
            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
                        
            sStr = sStr & "               Union All"
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, Birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, SEL6, CL_CLOSE, "
            sStr = sStr & "                      0 AS PAYOK, 0 AS PAYNOT, "
            sStr = sStr & "                      0 AS K_NUM, 0 AS M_NUM, 0 AS E_NUM, 0 AS N_NUM, 0 AS T_NUM, 0 AS TOT_NUM, "
            sStr = sStr & "                      SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS ZIP, ADDR1, ADDR2, TEL, CEL, "
            sStr = sStr & "                      TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI') AS REGDATE, GET_PAYGUBN(ORD_NO) AS PAYGBN, CASH_BILL_NUM, PRNT_TEL AS PRNT_TEL "
            sStr = sStr & "                 From CLSTD01TB"
            sStr = sStr & "                WHERE (PASS1 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS2 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS3 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS4 = '" & Trim(basModule.schcd) & "'" & " )"
            
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "              AND EXMTYPE = '1'"
            End If
            
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "              AND R_WAY IN ('1','3') "
            End If
            
            '<< 결재여부 >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "          AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "          AND EXMID IS NULL "
            End Select
            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
            
            '==========================================================================================================
            sStr = sStr & "               )"
            sStr = sStr & "        GROUP BY SCHNO"
            sStr = sStr & "       )"
            
            
        Case Else       ' /* 특정 합격차수의 합격자만 조회함 */
            sStr = sStr & " FROM (SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, Birth_ymd,"
            sStr = sStr & "              KAEYOL,"
            sStr = sStr & "              SEL1 , SEL2, SEL3, SEL4, SEL5,SEL6, CL_CLOSE, "
            sStr = sStr & "              0 AS PAYOK , 0 AS PAYNOT, "
            sStr = sStr & "              GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.schcd) & "') AS PAYTOT"     '< 전체집계 하는 함수
            sStr = sStr & "         From CLSTD01TB"
            sStr = sStr & "        WHERE PASS" & Trim(Right(cboPassCN, 30)) & " = '" & Trim(basModule.schcd) & "'"
            
            '>> 유/무시험 체크
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "      AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "      AND EXMTYPE = '1'"
            End If
            
            '>> 인터넷/학원
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
                sStr = sStr & "      AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
                sStr = sStr & "      AND R_WAY IN ('1','3') "
            End If
            
            '<< 결재여부 >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "  AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "  AND EXMID IS NULL "
            End Select
            sStr = sStr & "          AND CL_CLOSE IS NULL "
            
            sStr = sStr & "          AND BIGO2 IS NULL"                     '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
            
            sStr = sStr & "       )"
            
    End Select
    
    sStr = sStr & "   WHERE SCHNO > ' '"
    If Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) <> "" Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
        sStr = sStr & "               AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) = "" Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
        sStr = sStr & "               AND '99999'"
    ElseIf Trim(fpExmID_F.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) <> "" Then
        sStr = sStr & " AND EXMID BETWEEN '00000'"
        sStr = sStr & "               AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    Else
        'no action
    End If
    
    '<< 결재여부 >>
    Select Case Trim(Right(cboPay.Text, 30))
        Case "OK"
            sStr = sStr & " AND EXMID > ' ' "
        Case "NOT"
            sStr = sStr & " AND EXMID IS NULL "
    End Select
    
    If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then      ' 인문
        sStr = sStr & " AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
    End If
    
    If Trim(txtStdNM_F.Text) <> "" Then
        sStr = sStr & " AND STDNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
    End If
    If Trim(fpBirth_ymd_F.UnFmtText) <> "" Then
        sStr = sStr & " AND Birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
    End If
    If Trim(Right(cboSel1_SCH_F.Text, 30)) <> "X" Then
        sStr = sStr & " AND SEL1_SCH = '" & Trim(Right(cboSel1_SCH_F.Text, 30)) & "'"
    End If
    If Trim(Right(cboSel2_SCH_F.Text, 30)) <> "X" Then
        sStr = sStr & " AND SEL2_SCH = '" & Trim(Right(cboSel2_SCH_F.Text, 30)) & "'"
    End If
    
    sStr = sStr & "     AND CL_CLOSE IS NULL "
    
    sStr = sStr & "   ORDER BY EXMID "
    Text1.Text = sStr
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
        
    '>> 수험번호
'        If Trim(fpExmID_F.UnFmtText) > "" Then
'            sTmp = Trim(fpExmID_F.UnFmtText)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
'    '>> 학생명
'        If Trim(txtStdNM_F.Text) > "" Then
'            sTmp = "%" & Trim(txtStdNM_F.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> 주민번호
'        If Trim(fpBirth_ymd_F.UnFmtText) > "" Then
'            sTmp = "%" & Trim(fpBirth_ymd_F.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> 지망학원
'        If Trim(Right(cboSel1_SCH_F.Text, 30)) <> "X" Then
'            sTmp = Trim(Right(cboSel1_SCH_F.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("SEL1_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'        If Trim(Right(cboSel2_SCH_F.Text, 30)) <> "X" Then
'            sTmp = Trim(Right(cboSel2_SCH_F.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("SEL2_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    '2011-01-11 김한욱 언수외 및 총 합 전부 double 처리(노량진영향)
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
            
                If nRec = 1 Then        '< 인원수에 대한 부분은 한번만 출력하면 됩니다.
                    nTmp = 0:       If IsNumeric(.Fields("PAYOK")) = True Then nTmp = .Fields("PAYOK")
                        fpPayOK.value = nTmp
                        
                    nTmp = 0:       If IsNumeric(.Fields("PAYNOT")) = True Then nTmp = .Fields("PAYNOT")
                        fpPayNot.value = nTmp
                        
                    nTmp = 0:       If IsNumeric(.Fields("PAYTOT")) = True Then nTmp = .Fields("PAYTOT")
                        fpPayTot.value = nTmp
                End If
            
                sprSTD_F.MaxRows = sprSTD_F.MaxRows + 1
                sprSTD_F.Row = sprSTD_F.MaxRows
                
                
                sprSTD_F.Col = 1
                    sTmp = " ":  If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                    
                sprSTD_F.Col = 2
                    sTmp = " ":  If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD_F.Col = 3
                    sTmp = " ":  If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprSTD_F.Col = 4
                    sTmp = " ":  If IsNull(.Fields("SEL1_SCH")) = False Then sTmp = basCommonSTD.Get_SchName(.Fields("SEL1_SCH"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.Col = 5
                    sTmp = " ":  If IsNull(.Fields("SEL2_SCH")) = False Then sTmp = basCommonSTD.Get_SchName(.Fields("SEL2_SCH"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.Col = 6
                    sTmp = " ":  If IsNull(.Fields("Birth_ymd")) = False Then sTmp = Trim(.Fields("Birth_ymd"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                
                sprSTD_F.Col = 7
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = Trim(.Fields("K_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 8
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = Trim(.Fields("M_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 9
                    sTmp = "":   If IsNull(.Fields("SEL_X6")) = False Then sTmp = Trim(.Fields("SEL_X6"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = 10
                    nTmp = 0:   If IsNumeric(.Fields("T_NUM")) = True Then nTmp = Trim(.Fields("T_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 11
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = Trim(.Fields("E_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 12
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = Trim(.Fields("TOT_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("N_NUM")) = False Then sTmp = Trim(.Fields("N_NUM")) ' 내신등급
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "CENTER", LenB(sTmp), sTmp)
                    
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> 선택과목 (사탐/ 과탐)
                For ni = 1 To SATAM_COUNT Step 1
                
                    '파란색 세로 경게선 긋기
                    If ni Mod 4 = 1 Then: sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid

                    sprSTD_F.Col = sprSTD_F.Col + 1
                    
                    sGbn = "SEL" & Trim(CStr(ni))
                    sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                    If sTmp <> "" Then: sTmp = basGwamok.Get_StrGwaMokByCode(sTmp)   ' sTmp(코드)에 따른 과목이름얻어오기

                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)

                Next ni
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprSTD_F.Col = sprSTD_F.Col + 1
                If IsNull(.Fields("SEL_X2")) = True Then
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", 10, "")
                Else
                    If Trim(.Fields("SEL_X2")) = "00" Then
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", 10, "")
                    Else
                        Select Case Trim(.Fields("SEL_X2"))
                        
                            Case "31":  sTmp = "독어"
                            Case "32":  sTmp = "일어"
                            Case "33":  sTmp = "에스파냐어"
                            Case "34":  sTmp = "불어"
                            Case "35":  sTmp = "중국어"
                            Case "36":  sTmp = "한문"
                            
                            '<< 송파 >> : 2008.01.09
                            Case "37":  sTmp = "언어"
                            Case "38":  sTmp = "수리"
                            Case "39":  sTmp = "영어"
                            Case "40":  sTmp = "세계사"
                            Case "41":  sTmp = "세계지리"
                            
                            Case "81":  sTmp = "미적분"
                            Case "82":  sTmp = "이산수학"
                            Case "83":  sTmp = "확률통계"
                            Case "84":  sTmp = "수리나형"
                            
                        End Select
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> 논술
                For ni = 1 To 4 Step 1
                    sprSTD_F.Col = sprSTD_F.Col + 1
                    
                    sGbn = "SEL_N" & Trim(CStr(ni))
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        
                        If IsNull(.Fields(sGbn)) = False Then
                            If sTmp <> "" Then
                                Select Case sTmp
                                    Case "91":  sTmp = "언어"
                                    Case "92":  sTmp = "수리"
                                    Case "93":  sTmp = "외국어"         '< 변경
                                    Case "94":  sTmp = ""               '< 변경
                                    
                                End Select
                            End If
                            Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TEL")) = False Then sTmp = Trim(.Fields("TEL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("PRNT_TEL")) = False Then sTmp = Trim(.Fields("PRNT_TEL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("CEL")) = False Then sTmp = Trim(.Fields("CEL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("ZIP")) = False Then sTmp = Trim(.Fields("ZIP"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("ADDR1")) = False Then sTmp = Trim(.Fields("ADDR1"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("ADDR2")) = False Then sTmp = Trim(.Fields("ADDR2"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("REGDATE")) = False Then sTmp = Trim(.Fields("REGDATE"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("PAYGBN")) = False Then sTmp = Trim(.Fields("PAYGBN"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("CASH_BILL_NUM")) = False Then sTmp = Trim(.Fields("CASH_BILL_NUM"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":
                    If IsNull(.Fields("D_UNIVCD")) = False Then
                        sTmp = Trim(.Fields("D_UNIVCD"))
                    Else
                        sTmp = ""
                    End If
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":
                    If IsNull(.Fields("D_MAJORCD")) = False Then
                        sTmp = Trim(.Fields("D_MAJORCD"))
                    Else
                        sTmp = ""
                    End If
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                .MoveNext
            Next nRec
            
            sprSTD_F.Row = 1:       sprSTD_F.Row2 = sprSTD_F.MaxRows
            sprSTD_F.Col = 1:       sprSTD_F.Col2 = sprSTD_F.MaxCols
            sprSTD_F.BlockMode = True
                sprSTD_F.BackColor = basModule.BackColor1
                sprSTD_F.BackColorStyle = BackColorStyleUnderGrid
                
                sprSTD_F.Protect = True
                sprSTD_F.Lock = True
            sprSTD_F.BlockMode = False
            
            sprSTD_F.ColsFrozen = 3
            
        End If
    End With
    
    MsgBox "학생이 조회되었습니다.", vbInformation + vbOKOnly, "학생조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    MsgBox "학생조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "학생조회"
           
    On Error GoTo 0
End Sub

'>> 학생선택
Private Sub sprSTD_F_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprSTD_F
        If .ActiveRow < 1 Then Exit Sub
        
        Select Case KeyCode
            Case vbKeyUp, vbKeyDown, vbKeyNumpad8, vbKeyNumpad2
                .Enabled = False
                
                If .Tag = "" Then .Tag = "1"
                
                .Row = CLng(.Tag):  .Row2 = .Row
                
                
                .Col = 1:           .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.BackColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = .ActiveRow
                .Col = 1
                    Call Show_Select_STD(Trim(.Text))
                
                DoEvents
                
                .Row = .ActiveRow:  .Row2 = .Row
                .Col = 1:           .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Tag = Trim(CStr(.ActiveRow))
                
                .Enabled = True
                .SetFocus
                .SetActiveCell .ActiveCol, .ActiveRow
                
        End Select
    End With
End Sub

Private Sub sprSTD_F_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprSTD_F
        If .MaxRows < 1 Then Exit Sub
        
        sprSTD_F.Enabled = False
        
            If .Tag = "" Then .Tag = "1"
            
            .Row = CLng(.Tag):  .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Row = Row
            .Col = 1
                Call Show_Select_STD(Trim(.Text))
            
            DoEvents
            
            .Row = Row:         .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Tag = Trim(CStr(Row))
            
        sprSTD_F.Enabled = True
        sprSTD_F.SetFocus
        sprSTD_F.SetActiveCell Col, Row
        
    End With
    Call cboKaeyol_Click
End Sub

'>> 선택학생 보여주기
Private Sub Show_Select_STD(ByVal aSchNO As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim ni          As Integer
    Dim nLength     As Integer
    
    Dim sTmp        As String
    Dim sDiv()      As String
    Dim nDi         As Integer
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO   , ACID    , EXMID   , STDNM  ,D_UNIVCD, D_MAJORCD ,"
    sStr = sStr & " birth_ymd, "
    
    sStr = sStr & "         EXMTYPE , KAEYOL  ,"
    sStr = sStr & "         SEL1    , SEL2    , SEL3    , SEL4    , SEL5    ,SEL6, SEL7,"
    sStr = sStr & "         K_NUM   , M_NUM   , E_NUM   , N_NUM   , T_NUM   ,TOT_NUM ,"
    sStr = sStr & "         SEL1_SCH, SEL2_SCH,"
    sStr = sStr & "         PASS1 , PASS2, PASS3, PASS4 , TEL     , PRNT_TEL , CEL     , ORD_NO , "
    sStr = sStr & "         TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI') AS REGDATE, "
    sStr = sStr & "         GET_CANCELYN(SCHNO) AS WORKDAY, "
    sStr = sStr & "         GET_PAYGUBN(ORD_NO) AS PAYGBN, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2 , EMAIL   , PTS_SEL , MU_TYPE , SEX"
    sStr = sStr & "    From CLSTD01TB "
    sStr = sStr & "   WHERE SCHNO = '" & Trim(aSchNO) & "'"
    
    Select Case Trim(basModule.schcd)
        Case "W"
            sStr = sStr & "     AND (ACID  = '" & Trim(basModule.schcd) & "'"
            sStr = sStr & "          OR ACID = 'K')"
        Case "Q"
            sStr = sStr & "     AND (ACID  = '" & Trim(basModule.schcd) & "'"
            sStr = sStr & "          OR ACID = 'K')"
        Case Else
            'sStr = sStr & "     AND ACID  = '" & Trim(basModule.SchCD) & "'"
    End Select
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
    
'    '>> 학생코드
'        sTmp = Trim(aSchNO)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount <> 1 Then
            MsgBox "조회할 학생이 없습니다.", vbExclamation + vbOKOnly, "학생조회"
        Else
            .MoveFirst
            
            txtSchNo.Text = "":     If IsNull(.Fields("SCHNO")) = False Then txtSchNo.Text = Trim(.Fields("SCHNO"))
            fpExmID.Text = "":      If IsNull(.Fields("EXMID")) = False Then fpExmID.Text = Trim(.Fields("EXMID"))
            txtStdNM.Text = "":     If IsNull(.Fields("STDNM")) = False Then txtStdNM.Text = Trim(.Fields("STDNM"))
            
            txt_UNI.Text = "":     If IsNull(.Fields("D_UNIVCD")) = False Then txt_UNI.Text = Trim(.Fields("D_UNIVCD"))
            txt_MAJOR.Text = "":     If IsNull(.Fields("D_MAJORCD")) = False Then txt_MAJOR.Text = Trim(.Fields("D_MAJORCD"))
            
            fpBirth_ymd.Text = ""
            fpBirth_ymdS.Text = ""
                If IsNull(.Fields("Birth_ymd")) = False Then
                    fpBirth_ymd.Text = Trim(.Fields("Birth_ymd"))
                    fpBirth_ymdS.Text = Trim(.Fields("Birth_ymd"))
                End If
            
            txtTel.Text = "":       If IsNull(.Fields("TEL")) = False Then txtTel.Text = Trim(.Fields("TEL"))
            txt_P_Phone.Text = "":  If IsNull(.Fields("PRNT_TEL")) = False Then txt_P_Phone.Text = Trim(.Fields("PRNT_TEL"))
            txtCel.Text = "":       If IsNull(.Fields("CEL")) = False Then txtCel.Text = Trim(.Fields("CEL"))
            
            '< 2008.01.10 >
            fpZip.Text = "":        If IsNull(.Fields("ZIP")) = False Then fpZip.Text = Trim(.Fields("ZIP"))
            txtAddr1.Text = "":     If IsNull(.Fields("ADDR1")) = False Then txtAddr1.Text = Trim(.Fields("ADDR1"))
            txtAddr2.Text = "":     If IsNull(.Fields("ADDR2")) = False Then txtAddr2.Text = Trim(.Fields("ADDR2"))
            txtEmail.Text = "":     If IsNull(.Fields("EMAIL")) = False Then txtEmail.Text = Trim(.Fields("EMAIL"))
            
            
            txtOrdNo.Text = "":     If IsNull(.Fields("ORD_NO")) = False Then txtOrdNo.Text = Trim(.Fields("ORD_NO"))
            
            txtRegDate.Text = "":   If IsNull(.Fields("REGDATE")) = False Then txtRegDate.Text = Trim(.Fields("REGDATE"))
            txtPayGbn.Text = "":    If IsNull(.Fields("PAYGBN")) = False Then txtPayGbn.Text = Trim(.Fields("PAYGBN"))
            
            txtCancel.Text = "":    If IsNull(.Fields("WORKDAY")) = False Then txtCancel.Text = Trim(.Fields("WORKDAY"))
            
            '>> 성별 구분
            If IsNull(.Fields("SEX")) = False Then
                If Trim(.Fields("SEX")) = "M" Then
                    optSexMale.value = True
                Else
                    optSexFemale.value = True
                End If
            End If
            
            '유/무 시험
            optExmY.value = True
            If IsNull(.Fields("EXMTYPE")) = False Then
                If Trim(.Fields("EXMTYPE")) = "0" Then
                    optExmN.value = True
                Else
                    optExmY.value = True
                End If
            End If
            
            '>> 계열
            cboKaeyol.ListIndex = 0
            Call basCommonSTD.Set_CboKaeyol(cboKaeyol, basModule.schcd, Trim(.Fields("KAEYOL")))
            
            
            
            '>> 1지망학원
            If IsNull(.Fields("SEL1_SCH")) = False Then
                Call basCommonSTD.Set_CboSch(cboSel1_Sch, Trim(.Fields("SEL1_SCH")))
            End If
            
            
            '>> 2지망학원
            If IsNull(.Fields("SEL2_SCH")) = False Then
                Call basCommonSTD.Set_CboSch(cboSel2_Sch, Trim(.Fields("SEL2_SCH")))
            End If
            
           
            '>> 1차 합격학원
            If IsNull(.Fields("PASS1")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass1, Trim(.Fields("PASS1")))
            End If
            
            '>> 2차 합격학원
            If IsNull(.Fields("PASS2")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass2, Trim(.Fields("PASS2")))
            End If
            
            '>> 3차 합격학원
            If IsNull(.Fields("PASS3")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass3, Trim(.Fields("PASS3")))
            End If
                
            
            '>> 4차 합격학원
            If IsNull(.Fields("PASS4")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass4, Trim(.Fields("PASS4")))
            End If
            
            '2011-01-11 김한욱 언수외 및 총 합 전부 double 처리(노량진영향)
            fpK_Num.value = 0:  If IsNull(.Fields("K_NUM")) = False Then fpK_Num.value = CDbl(.Fields("K_NUM"))
            fpE_Num.value = 0:  If IsNull(.Fields("E_NUM")) = False Then fpE_Num.value = CDbl(.Fields("E_NUM"))
            fpM_Num.value = 0:  If IsNull(.Fields("M_NUM")) = False Then fpM_Num.value = CDbl(.Fields("M_NUM"))
            fpN_Num.value = 0:  If IsNull(.Fields("N_NUM")) = False Then fpN_Num.value = CDbl(.Fields("N_NUM")) ' 내신등급
            
            '2012-01-05 김한욱 언수외에 탐구 추가 총 합 전부 double 처리(노량진영향)
            cbo_Tamgoo.ListIndex = -1
            If IsNull(.Fields("SEL6")) = False Then
                Select Case Trim(Right(.Fields("SEL6"), 5))
                    '사탐
                    Case constSatamCodes(0)
                        cbo_Tamgoo.ListIndex = 1
                    Case constSatamCodes(1)
                        cbo_Tamgoo.ListIndex = 2
                    Case constSatamCodes(2)
                        cbo_Tamgoo.ListIndex = 3
                    Case constSatamCodes(3)
                        cbo_Tamgoo.ListIndex = 4
                    Case constSatamCodes(4)
                        cbo_Tamgoo.ListIndex = 5
                    Case constSatamCodes(5)
                        cbo_Tamgoo.ListIndex = 6
                    Case constSatamCodes(6)
                        cbo_Tamgoo.ListIndex = 7
                    Case constSatamCodes(7)
                        cbo_Tamgoo.ListIndex = 8
                    Case constSatamCodes(8)
                        cbo_Tamgoo.ListIndex = 9
                    Case constSatamCodes(9)
                        cbo_Tamgoo.ListIndex = 10
                    '과탐
                    Case "51"
                        cbo_Tamgoo.ListIndex = 12
                    Case "52"
                        cbo_Tamgoo.ListIndex = 13
                    Case "53"
                        cbo_Tamgoo.ListIndex = 14
                    Case "54"
                        cbo_Tamgoo.ListIndex = 15
                    Case "55"
                        cbo_Tamgoo.ListIndex = 16
                    Case "56"
                        cbo_Tamgoo.ListIndex = 17
                    Case "57"
                        cbo_Tamgoo.ListIndex = 18
                    Case "58"
                        cbo_Tamgoo.ListIndex = 19
                    Case Else
                        cbo_Tamgoo.ListIndex = 0
                End Select
            End If
            fpT_Num.value = 0:  If IsNull(.Fields("T_NUM")) = False Then fpT_Num.value = CDbl(.Fields("T_NUM"))
            
        '## 선택과목
            '>> 사탐
            '2011-05-17 김한욱 사회 부분 한과목 더 추가로 인한 수정 야간법의
            If (Trim(basModule.schcd) = "Q") Then
                For ni = 1 To SATAM_COUNT + 1 Step 1
                    chkSatam(ni).value = 0
                Next ni
            Else
                For ni = 1 To SATAM_COUNT Step 1
                    chkSatam(ni).value = 0
                Next ni
            End If
            
            If IsNull(.Fields("SEL1")) = False Then
                sTmp = Trim(.Fields("SEL1"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                '2011-05-17 김한욱  사회 부분 한과목 더 추가로 인한 수정 야간법의
                Dim arrIdx As Integer
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    
                    If sDiv(ni) <> "" Then
                        '현재 사탐의 코드 영역은 21~30까지. arrIdx = CInt(21) - (21-1)   , arrIdx = 1
                        arrIdx = CInt(sDiv(ni)) - SATAM_CLASS
                        If arrIdx > 0 And arrIdx <= chkSatam.UBound Then
                            chkSatam(arrIdx).value = 1
                        Else
                            MsgBox "DB의 사탐과목코드 값이 올바르지 않습니다. 다시 설정해주세요."
                        End If
                    End If
                Next ni
                
            End If
            
            '>> 제2외국어
            For ni = 1 To ENG2_COUNT Step 1
                chkEng2(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL2")) = False Then
                sTmp = Trim(.Fields("SEL2"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    If sDiv(ni) = "44" Then '베트남어 코드구함 :   베트남어코드:44 수리나형:43 이라서 +1해줌
                        chkEng2(CInt(sDiv(ni)) - 31).value = 1
                    Else
                        chkEng2(CInt(sDiv(ni)) - 30).value = 1
                    End If
                Next ni
            End If
            
            '>> 과탐
            For ni = 1 To 8 Step 1
                chkGwatam(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL3")) = False Then
                sTmp = Trim(.Fields("SEL3"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkGwatam(CInt(sDiv(ni)) - 50).value = 1
                Next ni
            End If
            '>> 수리
            For ni = 1 To 4 Step 1
                chkMath(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL4")) = False Then
                sTmp = Trim(.Fields("SEL4"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkMath(CInt(sDiv(ni)) - 80).value = 1
                Next ni
            End If
            '>> 논술
            For ni = 1 To 4 Step 1
                chkNonsul(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL5")) = False Then
                sTmp = Trim(.Fields("SEL5"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkNonsul(CInt(sDiv(ni)) - 90).value = 1
                Next ni
            End If
            
            '>> 클리닉 (노량진 요청)
            Call Init_Clinic(chkClinic_L, chkClinic_M, chkClinic_E)     ' 초기화
                        
            If IsNull(.Fields("SEL7")) = False Then
                Call Set_Clinic(chkClinic_L, chkClinic_M, chkClinic_E, .Fields("SEL7"))     ' 컨트롤 값 세팅
            End If
            
            
            '수리가/나 : 송파/ 송파마이맥
            Select Case Trim(basModule.schcd)
                Case "S", "N", "K"
                    If IsNull(.Fields("PTS_SEL")) = True Then
                        cboPTS_Sel.ListIndex = 2
                    Else
                        Select Case Trim(.Fields("PTS_SEL"))
                            Case "1"
                                cboPTS_Sel.ListIndex = 0
                            Case "2"
                                cboPTS_Sel.ListIndex = 1
                        End Select
                    End If
                    
                Case "P"
                    If IsNull(.Fields("PTS_SEL")) = True Then
                        cboPTS_Sel.ListIndex = 3
                    Else
                        Select Case Trim(.Fields("PTS_SEL"))
                            Case "8"
                                cboPTS_Sel.ListIndex = 0
                            Case "9"
                                cboPTS_Sel.ListIndex = 1
                            Case "6"
                                cboPTS_Sel.ListIndex = 2
                        End Select
                    End If
                    
                Case Else
                    cboPTS_Sel.ListIndex = 0
            End Select
            
            
            
            '수능등급
            If IsNull(.Fields("MU_TYPE")) = True Then
                cboMu_type.ListIndex = cboMu_type.ListCount - 1
            Else
                Call Set_Mu_type(cboMu_type, .Fields("MU_TYPE"))
            End If
            
           
            
            
        End If
    End With
    
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "선택학생 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
    
End Sub
















'## EXCEL 자료조회
Private Sub cmdGetExcel_Click()
    
    On Error GoTo ErrStmt
    
    cmdGetExcel.Enabled = False
        Call Get_Excel_Data
        
    cmdGetExcel.Enabled = True
    
    Exit Sub
ErrStmt:
    MsgBox "엑셀자료 가져오는 중 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생 엑셀자료 가져오기"
    On Error GoTo 0
    
End Sub

'##############################################################################
'################ 2012.11월 사용을 하지 않는것 같아서 폼에서 enable = false로 해둠.
'################            사용안함. 나중에 필요할경우 공통 모듈로 빼서 작업해야함.
'##############################################################################

Private Sub Get_Excel_Data()

'    Dim sPath       As String
'
'    ' Excel Data 처리
'    Dim xlsDBConn   As ADODB.Connection
'    Dim DBExCmd     As ADODB.Command
'    Dim DBExRec     As ADODB.Recordset
'
'    Dim sConn       As String
'    Dim sSql        As String
'
'    Dim nRow        As Long
'    Dim sTmp        As String
'    Dim nTmp        As Long
'
'    Dim nJumsu      As Long
'    Dim ni          As Long
'    Dim nC          As Long
'
'    On Error GoTo ErrStmt1
'
'    With dlgFile
'        .CancelError = True
'        .fileName = ""
'        .InitDir = App.Path
'        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
'        .DefaultExt = "*.XLS"
'        .ShowOpen
'
'        If (.fileName) = "" Then
'            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
'            Exit Sub
'        End If
'
'        sPath = .fileName
'
'    End With
'
'    On Error GoTo 0
'
'    On Error GoTo ErrStmt2                          '>> error 처리
'
'    Set xlsDBConn = New ADODB.Connection
'    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'            "Data Source=" & sPath & ";" & _
'            "Extended Properties=""Excel 8.0;HDR=no;"";"
'
'    With xlsDBConn
'        .ConnectionString = sConn                   ' 데이터베이스와 연결을 시도합니다.
'        .ConnectionTimeout = 30                     ' 제한 시간내에 연결이 되지 않으면 자동으로 끊습니다.
'        .Properties("Prompt") = adPromptNever       ' 이것은 ADO에서 기본 프롬프트 모드입니다.
'        .CursorLocation = adUseClient               ' 커서위치를 Client 쪽에 넣습니다.
'
'        .Open                                       ' 데이터베이스를 엽니다.
'
'        Do While .State And adStateConnecting
'            DoEvents
'        Loop
'    End With
'
''>> 엑셀 DB Open
'    sSql = ""
'    sSql = sSql & " SELECT * "
'    sSql = sSql & "   FROM [Sheet1$] "
'
'    Set DBExCmd = New ADODB.Command
'    Set DBExRec = New ADODB.Recordset
'
'    DBExCmd.ActiveConnection = xlsDBConn
'    DBExCmd.CommandText = sSql
'    DBExCmd.CommandType = adCmdText
'    DBExCmd.CommandTimeout = 30
'
'    DBExRec.Open DBExCmd, , adOpenStatic, adLockReadOnly, -1
'    Do While xlsDBConn.State And adStateExecuting
'        DoEvents
'    Loop
'
'    If DBExRec.RecordCount = 0 Then
'        Set DBExCmd = Nothing
'        Set DBExRec = Nothing
'        Set xlsDBConn = Nothing
'
'        MsgBox "Excel Data가 없습니다.", vbExclamation + vbOKOnly, "IT2008"
'        Exit Sub
'    End If
'
'
'    sprExcel_STD_Data.MaxRows = 0       ' 초기화
'
'
'    DBExRec.MoveFirst
'
'    '## header 1 line skip
'    DBExRec.MoveNext
'
'
'    For nRow = 2 To DBExRec.RecordCount Step 1
'    '학원코드
'        sTmp = "":  If IsNull(DBExRec.Fields(0)) = False Then sTmp = UCase(Trim(DBExRec.Fields(0)))
'        uExcel_StdData.ACID = sTmp
'    '수험번호
'        sTmp = "":  If IsNull(DBExRec.Fields(1)) = False Then sTmp = Trim(DBExRec.Fields(1))
'        uExcel_StdData.EXMID = sTmp
'    '학생명
'        sTmp = "":  If IsNull(DBExRec.Fields(2)) = False Then sTmp = Trim(DBExRec.Fields(2))
'        uExcel_StdData.STDNM = sTmp
'    '생년월일
'        sTmp = "":  If IsNull(DBExRec.Fields(3)) = False Then sTmp = Trim(DBExRec.Fields(3))
'        sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
'        If basFunction.LenKor(sTmp) > 6 Then
'            sTmp = Left(sTmp, 4) & "-" & Mid(sTmp, 5, 2) & "-" & Mid(sTmp, 7, 2)
'        End If
'        uExcel_StdData.Birth_ymd = sTmp
'    '유.무시험
'        sTmp = "1"
'        If IsNull(DBExRec.Fields(4)) = False Then
'            sTmp = UCase(Trim(DBExRec.Fields(4)))
'            Select Case sTmp
'                Case "0", "1"
'                    'no action
'                Case Else
'                    sTmp = "1"
'
'            End Select
'        End If
'        uExcel_StdData.EXMTYPE = sTmp
'    '계열
'        sTmp = "01"
'        If Trim(basModule.schcd) = "N" Then             '< 계열 : 2008.01.09 - 노량진
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'                    Case "예체", "예"
'                        sTmp = "03"
'
'                    Case "수리(나)", "수리나"
'                        sTmp = "04"
'                    Case "인문수능", "수능인문"
'                        sTmp = "05"
'                    Case "자연수능", "수능자연"
'                        sTmp = "06"
'
'                    Case "신설인문"
'                        sTmp = "07"
'                    Case "신설자연"
'                        sTmp = "08"
'                    'Case "신설수능인문"
'                    '    sTmp = "09"
'                    'Case "신설수능자연"
'                    '    sTmp = "10"
'
'                    Case "편입인문", "편인"
'                        sTmp = "11"
'                    Case "편입자연", "편자"
'                        sTmp = "12"
'                    Case "편예체", "편예"
'                        sTmp = "13"
'
'                    Case "편수리(나)", "편수리나"
'                        sTmp = "14"
'                    Case "편인문수능", "편수능인문"
'                        sTmp = "15"
'                    Case "편자연수능", "편수능자연"
'                        sTmp = "16"
'                    Case "서울대인문"
'                        sTmp = "21"
'                    Case "서울대자연"
'                        sTmp = "22"
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then           '< 계열 : 2008.01.09 - 노량진, 2008.03.24 - 강남
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'
'                    Case "주간법대", "주법"
'                        sTmp = "04"
'                    Case "주간의대", "주의"
'                        sTmp = "05"
'
'                    Case "야간법대", "야법"
'                        sTmp = "06"
'                    Case "야간의대", "야의"
'                        sTmp = "07"
'
'                    Case "선착순인문"
'                        sTmp = "11"
'                    Case "선착순자연"
'                        sTmp = "12"
'
'                    Case "선착순인문16"
'                        sTmp = "16"
'                    Case "선착순자연17"
'                        sTmp = "17"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "S" Then             '< 계열 : 2008.02.15 - 송파
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'
'                    Case "특인", "특별인문"
'                        sTmp = "03"
'                    Case "특자", "특별자연"
'                        sTmp = "04"
'
'                    Case "신설수능인문"
'                        sTmp = "05"
'                    Case "신설수능자연"
'                        sTmp = "06"
'
'                    Case "신설인문"
'                        sTmp = "11"
'                    Case "신설자연"
'                        sTmp = "12"
'
'                    Case "신설수리나형"
'                        sTmp = "08"
'
'                    Case "인문프리미엄"
'                        sTmp = "18"
'                    Case "자연프리미엄"
'                        sTmp = "19"
'
'                    Case "야간서울대인문"
'                        sTmp = "23"
'                    Case "야간서울대자연"
'                        sTmp = "24"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "P" Then             '< 계열 : 2008.02.15 - 마송
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'
'                    Case "특인", "특별인문"
'                        sTmp = "03"
'                    Case "특자", "특별자연"
'                        sTmp = "04"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'
'        ElseIf Trim(basModule.schcd) = "J" Then             '< 양재
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'
'                    Case "신설인문"
'                        sTmp = "11"
'                    Case "신설자연"
'                        sTmp = "12"
'
'                    Case "인문프리미엄"
'                        sTmp = "18"
'                    Case "자연프리미엄"
'                        sTmp = "19"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "B" Then             '< 양재
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'                    Case "예체", "예"
'                        sTmp = "03"
'                    Case "인문PS반"
'                        sTmp = "23"
'                    Case "자연PM반"
'                        sTmp = "24"
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        Else
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "인문", "인"
'                        sTmp = "01"
'                    Case "자연", "자"
'                        sTmp = "02"
'                    Case "예체", "예"
'                        sTmp = "03"
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        End If
'        uExcel_StdData.kaeyol = sTmp
'
'    '1 지망학원
'        sTmp = Trim(basModule.schcd)
'        If IsNull(DBExRec.Fields(6)) = False Then
'            sTmp = UCase(Trim(DBExRec.Fields(6)))
'            Select Case sTmp
'                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
'                    ' NEXT
'                Case "노량진"
'                    sTmp = "N"
'                Case "강남"
'                    sTmp = "K"
'                Case "송파"
'                    sTmp = "S"
'                Case "송파M", "송파마이맥", "송파 MIMAC", "송파MIMAC", "마송"
'                    sTmp = "P"
'                Case "강남M", "강남마이맥", "강남 MIMAC", "강남MIMAC", "마강"
'                    sTmp = "M"
'
'                Case "주말법의대", "주말법", "주법"
'                    sTmp = "W"
'                Case "야간법의대", "야간법", "야법"
'                    sTmp = "Q"
'
'                Case "양재"
'                    sTmp = "J"
'
'                Case "부산"
'                    sTmp = "B"
'
'                Case Else
'                    sTmp = Trim(basModule.schcd)
'            End Select
'        End If
'        uExcel_StdData.WANT_ACID1 = sTmp
'
'    '2 지망학원
'        sTmp = Trim(basModule.schcd)
'        If IsNull(DBExRec.Fields(7)) = False Then
'            sTmp = UCase(Trim(DBExRec.Fields(7)))
'            Select Case sTmp
'                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
'                    ' NEXT
'                Case "노량진"
'                    sTmp = "N"
'                Case "강남"
'                    sTmp = "K"
'                Case "송파"
'                    sTmp = "S"
'                Case "송파M", "송파마이맥", "송파 MIMAC", "송파MIMAC", "마송"
'                    sTmp = "P"
'                Case "강남M", "강남마이맥", "강남 MIMAC", "강남MIMAC", "마강"
'                    sTmp = "M"
'
'                Case "주말법의대", "주말법", "주법"
'                    sTmp = "W"
'                Case "야간법의대", "야간법", "야법"
'                    sTmp = "Q"
'
'                Case "양재"
'                    sTmp = "J"
'
'                Case "부산"
'                    sTmp = "B"
'
'                Case Else
'                    sTmp = Trim(basModule.schcd)
'            End Select
'        End If
'        uExcel_StdData.WANT_ACID2 = sTmp
'
'    '국어
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(8)) = True Then nTmp = CLng(Trim(DBExRec.Fields(8)))
'        uExcel_StdData.KOR = nTmp
'    '영어
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(9)) = True Then nTmp = CLng(Trim(DBExRec.Fields(9)))
'        uExcel_StdData.ENG = nTmp
'    '수학
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(10)) = True Then nTmp = CLng(Trim(DBExRec.Fields(10)))
'        uExcel_StdData.MAT = nTmp
'
'    '사탐
'        uExcel_StdData.SATAM1 = ""
'        uExcel_StdData.SATAM2 = ""
'        uExcel_StdData.SATAM3 = ""
'        uExcel_StdData.SATAM4 = ""
'        uExcel_StdData.SATAM5 = ""
'        uExcel_StdData.SATAM6 = ""
'        uExcel_StdData.SATAM7 = ""
'        uExcel_StdData.SATAM8 = ""
'        uExcel_StdData.SATAM9 = ""
'        uExcel_StdData.SATAM10 = ""
'
'
'        For ni = 1 To SATAM_COUNT Step 1
'            sTmp = ""
'            nC = 10 + ni
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case constSatams(0)
'                    uExcel_StdData.SATAM1 = constSatamCodes(0) & "|"
'                Case constSatams(1)
'                    uExcel_StdData.SATAM2 = constSatamCodes(1) & "|"
'                Case constSatams(2)
'                    uExcel_StdData.SATAM3 = constSatamCodes(2) & "|"
'                Case constSatams(3)
'                    uExcel_StdData.SATAM4 = constSatamCodes(3) & "|"
'                Case constSatams(4)
'                    uExcel_StdData.SATAM5 = constSatamCodes(4) & "|"
'                Case constSatams(5)
'                    uExcel_StdData.SATAM6 = constSatamCodes(5) & "|"
'                Case constSatams(6)
'                    uExcel_StdData.SATAM7 = constSatamCodes(6) & "|"
'                Case constSatams(7)
'                    uExcel_StdData.SATAM8 = constSatamCodes(7) & "|"
'                Case constSatams(8)
'                    uExcel_StdData.SATAM9 = constSatamCodes(8) & "|"
'                Case constSatams(9)
'                    uExcel_StdData.SATAM10 = constSatamCodes(9) & "|"
'            End Select
'        Next ni
'    '제2외국어
'        uExcel_StdData.ENG2 = ""
'
'        sTmp = ""
'            nC = 10 + 11 + 1
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "독어"
'                    uExcel_StdData.ENG2 = "31|"
'                Case "일어"
'                    uExcel_StdData.ENG2 = "32|"
'                Case "에파", "에스파냐"
'                    uExcel_StdData.ENG2 = "33|"
'                Case "불어"
'                    uExcel_StdData.ENG2 = "34|"
'                Case "중국", "중어"
'                    uExcel_StdData.ENG2 = "35|"
'                Case "한문"
'                    uExcel_StdData.ENG2 = "36|"
'
'                '<< 송파 >> : 2008.01.09
'                Case "언어"
'                    uExcel_StdData.ENG2 = "37|"
'                Case "수리"
'                    uExcel_StdData.ENG2 = "38|"
'                Case "영어"
'                    uExcel_StdData.ENG2 = "39|"
'                Case "세계사"
'                    uExcel_StdData.ENG2 = "40|"
'                Case "세계지리"
'                    uExcel_StdData.ENG2 = "41|"
'                Case "아랍어"
'                    uExcel_StdData.ENG2 = "42|"
'
'            End Select
'    '과탐
'        uExcel_StdData.GWATAM1 = ""
'        uExcel_StdData.GWATAM2 = ""
'        uExcel_StdData.GWATAM3 = ""
'        uExcel_StdData.GWATAM4 = ""
'        uExcel_StdData.GWATAM5 = ""
'        uExcel_StdData.GWATAM6 = ""
'        uExcel_StdData.GWATAM7 = ""
'        uExcel_StdData.GWATAM8 = ""
'
'        For ni = 1 To 8 Step 1
'            sTmp = ""
'            nC = 10 + ni
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "물1"
'                    uExcel_StdData.GWATAM1 = "51|"
'                Case "화1"
'                    uExcel_StdData.GWATAM2 = "52|"
'                Case "생1"
'                    uExcel_StdData.GWATAM3 = "53|"
'                Case "지1"
'                    uExcel_StdData.GWATAM4 = "54|"
'                Case "물2"
'                    uExcel_StdData.GWATAM5 = "55|"
'                Case "화2"
'                    uExcel_StdData.GWATAM6 = "56|"
'                Case "생2"
'                    uExcel_StdData.GWATAM7 = "57|"
'                Case "지2"
'                    uExcel_StdData.GWATAM8 = "58|"
'            End Select
'        Next ni
'    '수리
'        uExcel_StdData.SURI = ""
'
'        sTmp = ""
'            nC = 10 + 11 + 1
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "미적"
'                    uExcel_StdData.SURI = "81|"
'                Case "이산"
'                    uExcel_StdData.SURI = "82|"
'                Case "확률"
'                    uExcel_StdData.SURI = "83|"
'                Case "나형"
'                    uExcel_StdData.SURI = "84|"
'            End Select
'    '논술
'        uExcel_StdData.NONSUL1 = ""
'        uExcel_StdData.NONSUL2 = ""
'        uExcel_StdData.NONSUL3 = ""
'        uExcel_StdData.NONSUL4 = ""
'
'        For ni = 1 To 4 Step 1
'            sTmp = ""
'            nC = 10 + 11 + 1 + ni
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "언어"
'                    uExcel_StdData.NONSUL1 = "91|"
'                Case "수리"
'                    uExcel_StdData.NONSUL2 = "92|"
'                Case "외국어"
'                    uExcel_StdData.NONSUL3 = "93|"
'                Case ""
'                    uExcel_StdData.NONSUL4 = "94|"
'            End Select
'        Next ni
'
'
'    '## 스프레드에 데이터 넣기 --------------------------------------------------------------------
'        With sprExcel_STD_Data
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows:            .RowHeight(.Row) = 13
'
'            '>> 학원
'                .Col = 1
'                    sTmp = uExcel_StdData.ACID
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 수험번호
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.EXMID
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 학생명
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.STDNM
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 생년월일
'                .Col = .Col + 1
'                    sTmp = Replace(uExcel_StdData.Birth_ymd, "-", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 유.무시험
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.EXMTYPE
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 계열
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.kaeyol
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 1 지망학원
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.WANT_ACID1
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 2 지망학원
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.WANT_ACID2
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 국어
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.KOR
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'            '>> 영어
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.ENG
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'            '>> 수학
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.MAT
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'
'            '>> 사탐
'                .Col = .Col + 1
'                    sTmp = ""
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM1)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM2)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM3)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM4)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM5)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM6)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM7)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM8)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM9)
'                    sTmp = sTmp & Trim(uExcel_StdData.SATAM10)
'
'                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 제2외국어
'                .Col = .Col + 1
'                    sTmp = Trim(uExcel_StdData.ENG2)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 과탐
'                .Col = .Col + 1
'                    sTmp = ""
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM1)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM2)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM3)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM4)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM5)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM6)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM7)
'                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM8)
'
'                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 수리
'                .Col = .Col + 1
'                    sTmp = Trim(uExcel_StdData.SURI)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 논술
'                .Col = .Col + 1
'                    sTmp = ""
'                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL1)
'                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL2)
'                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL3)
'                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL4)
'
'                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'        End With
'
'        DBExRec.MoveNext
'
'    Next nRow
'
'
'
'    With sprExcel_STD_Data
'        If .MaxRows > 0 Then
'            .Row = 1:   .Row2 = .MaxRows
'            .Col = 1:   .Col2 = .MaxCols
'            .BlockMode = True
'                .BackColor = basModule.BackColor1
'                .BackColorStyle = BackColorStyleUnderGrid
'            .BlockMode = False
'
'            '.ColsFrozen = 3
'            '.SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
'
'        End If
'    End With
'
'
'    Set DBExRec = Nothing
'    Set DBExCmd = Nothing
'    Set xlsDBConn = Nothing
'
'    MsgBox "학생 엑셀자료를 가지고 왔습니다.", vbInformation + vbOKOnly, Me.Caption
'
'    On Error GoTo 0
'    Exit Sub
'ErrStmt1:
'    MsgBox "엑셀 파일선택을 하십시요.", vbExclamation + vbOKOnly, Me.Caption
'    Exit Sub
'ErrStmt2:
'    Set DBExRec = Nothing
'    Set DBExCmd = Nothing
'    xlsDBConn.Close
'    Set xlsDBConn = Nothing
'
'    MsgBox "EXCEL 자료 Open시 에러가 발생하였습니다.", vbCritical + vbOKOnly, Me.Caption
'    On Error GoTo 0
'    Exit Sub
End Sub

Private Sub sprExcel_STD_Data_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprExcel_STD_Data
        If .MaxRows < 1 Then Exit Sub
        
        sprExcel_STD_Data.Enabled = False
        
            If .Tag = "" Then .Tag = "1"
            
            .Row = CLng(.Tag):  .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Row = Row:         .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Tag = Trim(CStr(Row))
            
        sprExcel_STD_Data.Enabled = True
        
        sprExcel_STD_Data.SetFocus
        sprExcel_STD_Data.SetActiveCell Col, Row
        
    End With
    
End Sub


Private Sub sprExcel_STD_Data_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprExcel_STD_Data
        If .ActiveRow < 1 Then Exit Sub
        
        Select Case KeyCode
            Case vbKeyUp, vbKeyDown, vbKeyNumpad8, vbKeyNumpad2
                .Enabled = False
                
                If .Tag = "" Then .Tag = "1"
                
                .Row = CLng(.Tag):  .Row2 = .Row
                
                
                .Col = 1:           .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.BackColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = .ActiveRow:  .Row2 = .Row
                .Col = 1:           .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Tag = Trim(CStr(.ActiveRow))
                
                .Enabled = True
                .SetFocus
                .SetActiveCell .ActiveCol, .ActiveRow
                
        End Select
    End With
End Sub





'>> 학생등록
Private Sub cmdExcelSave_Click()
    Dim bRet        As Boolean
    
    '>> 체크조건
    If sprExcel_STD_Data.MaxRows = 0 Then
        MsgBox "등록할 학생이 없습니다.", vbExclamation + vbOKOnly, "엑셀로 학생등록"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdExcelSave.Enabled = False
        bRet = Save_Excel_Stdin             '<< 학생등록
            
    cmdExcelSave.Enabled = True
            
    If bRet = True Then
        MsgBox "학생 엑셀자료로 등록하였습니다.", vbInformation + vbOKOnly, "엑셀로 학생등록"
    Else
        MsgBox "학생 엑셀자료 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "엑셀로 학생등록"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "학생 엑셀자료 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "엑셀로 학생등록"
    On Error GoTo 0
    
End Sub

'>> 학생등록
Private Function Save_Excel_Stdin() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nRow        As Long
    Dim nTotJumsu   As Long
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    For nRow = 1 To sprExcel_STD_Data.MaxRows Step 1
        
        sprExcel_STD_Data.Row = nRow
    
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
        '>> 등록 신규
            sTmp = "INSERT"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_STYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 시스템코드
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 학원코드
            sprExcel_STD_Data.Col = 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        '>> 수험번호
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 학생명
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 생년월일
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text)):     sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_BIRTH_YMD", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        '>> 유/무시험 구분
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_EXMTYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 계열
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_KAEYOL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        '## 선택과목 ###
            '>> 사탐과목 선택
            sprExcel_STD_Data.Col = 12
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> 제2외국어 선택
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> 과탐과목 선택
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> 수리과목 선택
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> 논술과목 선택
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL5", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        nTotJumsu = 0

        '>> 국어점수
            sprExcel_STD_Data.Col = 9
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_K_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> 영어점수
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_E_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> 수학점수
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_M_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> 합계
            nTmp = nTotJumsu
                Set DBParam = DBCmd.CreateParameter("V_TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                
        '>> 내신등급
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
            If Trim(sprExcel_STD_Data.Text) > " " Then
                nTmp = CLng(Trim(sprExcel_STD_Data.Text))
            Else
                nTmp = 0
            End If
            
            Set DBParam = DBCmd.CreateParameter("V_N_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

        '>> 1지망 학원
            sprExcel_STD_Data.Col = 7
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL1_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 2지망 학원
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL2_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        '>> 1지망 합격학원
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 2지망 합격학원
            sTmp = ""

            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 3지망 합격학원
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 4지망 합격학원
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                
        '>> 전화번호
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 주소
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_CEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        '>> 데이터 등록
        DBCmd.CommandType = adCmdStoredProc
        DBCmd.CommandText = "PG_STD.PROC_STD_SAVE"
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute
        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
    
    Next nRow
    
    
    Save_Excel_Stdin = True
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_Excel_Stdin = False
    
End Function










'## 전체학생 데이터 받기
Private Sub cmdAllStdData_Click()
'    Dim DBCmd           As ADODB.Command
'    Dim DBRec           As ADODB.Recordset
'    Dim DBParam         As ADODB.Parameter
'
'    Dim nLength         As Long
'    Dim sStr            As String
'    Dim ni              As Integer
'
'    Dim nRec            As Long
'
'
'    Dim sTmp            As String
'    Dim nTmp            As Long
'    Dim nRet            As Long
'
'    Dim sExcelFileName  As String
'    Dim sExcelLogFile   As String
'
    Call imgExcel_Click
    Exit Sub
'
'    '> 초기화
'    sprStdData.MaxRows = 0
'
'    On Error GoTo ErrStmt1
'
'    With dlgFile
'        .CancelError = True
'        .fileName = ""
'        .InitDir = App.Path
'        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
'        .DefaultExt = "*.XLS"
'        .ShowSave
'
'        If (.fileName) = "" Then
'            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
'            Exit Sub
'        End If
'
'        sExcelFileName = .fileName
'
'        ni = InStrRev(sExcelFileName, "\", -1, vbTextCompare)
'        sExcelLogFile = Mid(sExcelFileName, 1, ni) & "\" & Mid(sExcelFileName, ni + 1, Len(sExcelFileName) - ni + 1 - 5)
'
'    End With
'
'    On Error GoTo 0
'
'    On Error GoTo ErrStmt
'
'    sStr = ""
'    sStr = sStr & "  SELECT SCHNO AS 시스템코드   , "
'    sStr = sStr & "         ACID  AS 학원   , "
'    sStr = sStr & "         EXMID AS 수험번호, STDNM AS 학생, D_UNIVCD AS 지원대학, D_MAJORCD AS 지원단대,"
'
'    sStr = sStr & " Birth_ymd ,"
'
'    sStr = sStr & "         DECODE(EXMTYPE,'0','무시험','1','유시험') AS 시험형태, "
'    sStr = sStr & "         DECODE(KAEYOL,'01','인문',"
'    sStr = sStr & "                       '02','자연',"
''<< 계열 >> : 2008.01.09
'    If Trim(basModule.SchCD) = "N" Then
'        sStr = sStr & "                   '03','예체',"
'        sStr = sStr & "                   '04','수리(나)',"
'        sStr = sStr & "                   '05','인문수능',"
'        sStr = sStr & "                   '06','자연수능',"
'
'        sStr = sStr & "                   '07','신설인문',"
'        sStr = sStr & "                   '08','신설자연',"
'        sStr = sStr & "                   '09','신설수능인문',"
'        sStr = sStr & "                   '10','신설수능자연',"
'
'        sStr = sStr & "                   '11','편)인문',"
'        sStr = sStr & "                   '12','편)자연',"
'        sStr = sStr & "                   '13','편)예체',"
'        sStr = sStr & "                   '14','편)수리(나)',"
'        sStr = sStr & "                   '15','편)인문수능',"
'        sStr = sStr & "                   '16','편)자연수능',"
'        sStr = sStr & "                   '21','서울대인문',"
'        sStr = sStr & "                   '22','서울대자연',"
'    End If
''<< 계열 >> : 2008.01.09
'    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
'        sStr = sStr & "                   '04','주말법대',"
'        sStr = sStr & "                   '05','주말의대',"
'        sStr = sStr & "                   '06','야간법대',"
'        sStr = sStr & "                   '07','야간의대',"
'
'        sStr = sStr & "                   '11','선착순인문',"
'        sStr = sStr & "                   '12','선착순자연',"
'
'        sStr = sStr & "                   '16','선착순인문16',"
'        sStr = sStr & "                   '17','선착순자연17',"
'    End If
''<< 계열 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "S" Then
'        sStr = sStr & "                   '03','예체능',"
'        'sStr = sStr & "                   '04','특별자연',"
'
'        sStr = sStr & "                   '05','수능인문',"
'        sStr = sStr & "                   '06','수능자연',"
'
'        sStr = sStr & "                   '11','신설인문',"
'        sStr = sStr & "                   '12','신설자연',"
'
'        sStr = sStr & "                   '18','인문프리미엄',"
'        sStr = sStr & "                   '19','자연프리미엄',"
'        sStr = sStr & "                   '21','서울대특별인문',"
'        sStr = sStr & "                   '22','서울대특별자연',"
'        sStr = sStr & "                   '23','야간서울대인문',"
'        sStr = sStr & "                   '24','야간서울대자연',"
'    End If
''<< 계열 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "P" Then         '< 마송
'        sStr = sStr & "                   '03','특별인문',"
'        sStr = sStr & "                   '04','특별자연',"
'    End If
'
'    If Trim(basModule.SchCD) = "J" Then         '< 양재
'        sStr = sStr & "                   '11','신설인문',"
'        sStr = sStr & "                   '12','신설자연',"
'
'        sStr = sStr & "                   '18','인문프리미엄',"
'        sStr = sStr & "                   '19','자연프리미엄',"
'        sStr = sStr & "                   '21','서울대특별인문',"
'        sStr = sStr & "                   '22','서울대특별자연',"
'
'    End If
'
''<< 계열 >> : 2009.01.09
'    If Trim(basModule.SchCD) = "B" Then         '< 부산
'        sStr = sStr & "                   '05','특별인문',"
'        sStr = sStr & "                   '06','특별자연',"
'        sStr = sStr & "                   '07','연고대인문',"
'        sStr = sStr & "                   '08','연고대자연',"
'        sStr = sStr & "                   '09','심화인문',"
'        sStr = sStr & "                   '10','심화자연',"
'        sStr = sStr & "                   '23','인문PS반',"
'        sStr = sStr & "                   '24','자연PM반',"
'    End If
'
'    sStr = sStr & "                       '','기타') AS 계열,"
'
'    sStr = sStr & "     /* 사탐, 과탐 분리 */"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* 사탐-한국사 */"
'    sStr = sStr & "             '" & constSatamCodes(0) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
'    sStr = sStr & "             '51'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL1,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* 사탐-세계사 */"
'    sStr = sStr & "             '" & constSatamCodes(1) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
'    sStr = sStr & "             '52'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL2,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* 사탐-동아시아사 */"
'    sStr = sStr & "             '" & constSatamCodes(2) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생명과학1 */"
'    sStr = sStr & "             '53'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL3,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* 사탐-한국지리 */"
'    sStr = sStr & "             '" & constSatamCodes(3) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
'    sStr = sStr & "             '54'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL4,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* 사탐-세계지리 */"
'    sStr = sStr & "             '" & constSatamCodes(4) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
'    sStr = sStr & "             '55'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL5,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* 사탐-생활과윤리 */"
'    sStr = sStr & "             '" & constSatamCodes(5) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
'    sStr = sStr & "             '56'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL6,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* 사탐-윤리와사상 */"
'    sStr = sStr & "             '" & constSatamCodes(6) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생명과학2 */"
'    sStr = sStr & "             '57'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL7,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* 사탐-법과정치 */"
'    sStr = sStr & "             '" & constSatamCodes(7) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
'    sStr = sStr & "             '58'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL8,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* 사탐-경제 */"
'    sStr = sStr & "             '" & constSatamCodes(8) & "'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END SEL9,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* 사탐-사회문화 */"
'    sStr = sStr & "             '" & constSatamCodes(9) & "'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END SEL10,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* 제2외국어 & 수리 */"
'    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에파'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
'    '<< 송파 >> : 2008.01.09
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '언어'"
'
'    Select Case Trim(basModule.SchCD)
'                Case "S"
'                    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '논술'"
'                Case Else
'                    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '수리'"
'    End Select
'
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '영어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '세계사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '세지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '아랍어'"
'
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '나형'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END END END END END END END END END END END END END END END 제2선택,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* 논술 */"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
'    sStr = sStr & "             '언어'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 언어논술,"
'
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
'
'    Select Case Trim(basModule.SchCD)
'                Case "S"
'                    sStr = sStr & "             '언수외'"
'                Case Else
'                    sStr = sStr & "             '수리'"
'    End Select
'
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'
'    Select Case Trim(basModule.SchCD)
'                Case "S"
'                    sStr = sStr & "         END 사탐선택,"
'                Case Else
'                    sStr = sStr & "         END 수리논술,"
'    End Select
'
'    Select Case Trim(basModule.SchCD)
'                Case "S"
'                    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
'                    sStr = sStr & "             '논술'"
'                Case Else
'                    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
'                    sStr = sStr & "             '외국어'"
'    End Select
'
'
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 사탐논술,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< 변경
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 과탐논술,"
'    sStr = sStr & "  "
'    sStr = sStr & "         CY_ACNT AS 가상계좌수강료, CY_ACNT2 AS 가상계좌교재비, CY_ACNT3 AS 가상계좌기타, TOT_AMT AS 전체금액    ,"
'    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS 기본금액1  ,"
'    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS 기본금액2  ,"
'    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS 기본금액3  ,"
'    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS 기본금액4  ,"
'    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS 기본금액5  ,"
'    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS 기본금액6  ,"
'    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS 기본금액7  ,"
'    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS 기본금액8  ,"
'    sStr = sStr & "         NVL(BASE_AMT9    ,0) AS 기본금액9  ,"
'    sStr = sStr & "         NVL(BASE_AMT10   ,0) AS 기본금액10 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS 탐구영역금액1 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS 탐구영역금액2 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS 탐구영역금액3 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS 탐구영역금액4 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS 탐구영역금액5 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS 탐구영역금액6 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS 탐구영역금액7 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS 탐구영역금액8 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS 탐구영역금액9 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS 탐구영역금액10,"
'    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS 탐구영역금액11,"
'    sStr = sStr & "         NVL(TAMGU_AMT12  ,0) AS 탐구영역금액12,"
'
'    sStr = sStr & "         K_NUM AS 언어점수, M_NUM AS 수학점수, E_NUM AS 영어점수, T_NUM AS 탐구점수,"
'    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)+NVL(T_NUM,0)) AS 전체점수, N_NUM AS 내신등급,"
'
'
'    sStr = sStr & "         DECODE(SEL1_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 제1지망,"
'    sStr = sStr & "         DECODE(SEL2_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 제2지망,"
'
'    sStr = sStr & "         DECODE(PASS1,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 합격1   ,"
'    sStr = sStr & "         DECODE(PASS2,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 합격2   ,"
'    sStr = sStr & "         DECODE(PASS3,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 합격3   ,"
'    sStr = sStr & "         DECODE(PASS4,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','H','양재', 'B','부산') AS 합격4   ,"
'
'
'    sStr = sStr & "         DECODE(SEX,'M','남','F','여') AS 성별        , "
'    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS 우편번호, ADDR1 AS 우편주소      , ADDR2 AS 상세주소     ,"
'    sStr = sStr & "         TEL AS 전화번호, CEL AS 핸드폰        , EMAIL AS 이메일     ,"
'    sStr = sStr & "         HIGH_SCH AS 고등학교 , GRADE_YEAR AS 졸업년도 ,"
'    sStr = sStr & "         PRNT_NM AS 학부모명 , DECODE(PRNT_RLTN,'1','부','2','모','3','기타') AS 학부모관계, "
'    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS 학부모_우편번호, PRNT_ADDR1 AS 학부모_우편주소 , PRNT_ADDR2 AS 학부모_상세주소,"
'    sStr = sStr & "         PRNT_TEL AS 학부모_전화번호  , PRNT_CEL AS 학부모_핸드폰   , PRNT_JOB AS 학부모_직업   , PRNT_W_TEL AS 학부모_직장전화 ,"
'    sStr = sStr & "         PHOTO_PATH AS 사진저장장소, "
'    sStr = sStr & "         DECODE(R_WAY,'1','학원등록','2','인터넷등록','3','학원등록') AS 등록번호, "
'    sStr = sStr & "         ORD_NO AS 주문번호, "
'    sStr = sStr & "         ACID||EXMID AS 이미지파일명, "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
'    sStr = sStr & "         REGDATE AS 등록일자, GET_PAYGUBN(ORD_NO) AS 결재방법, CASH_BILL_NUM AS 현금영수증,"
'    sStr = sStr & "         DECODE(MU_TYPE,'1','수능평가','2','6월 평가원','3','9월 평가원','4','6월 평가원','5','9','내신등급','9월 평가원','') AS 등급, "
'    sStr = sStr & "         CL_CLOSE AS 완료년월  "
'
'    sStr = sStr & "    FROM ( "
'
'            sStr = sStr & "  SELECT SCHNO           ,"
'            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
'            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
'            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
'            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
'            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
'            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
'            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
'            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
'            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(T_NUM     ) AS T_NUM     , MAX(N_NUM     ) AS N_NUM      , MAX(TOT_NUM   ) AS TOT_NUM   ,"
'            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
'            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
'            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(CY_ACNT2  ) AS CY_ACNT2   , MAX(CY_ACNT3  ) AS CY_ACNT3  , MAX(TOT_AMT   ) AS TOT_AMT    ,"
'            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
'            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
'            sStr = sStr & "         MAX(BASE_AMT9 ) AS BASE_AMT9  , MAX(BASE_AMT10) AS BASE_AMT10 , "
'            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
'            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11, MAX(TAMGU_AMT12) AS TAMGU_AMT12, "
'            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
'            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
'            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
'            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
'            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
'            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
'            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
'            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM, "
'            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(MU_TYPE   ) AS MU_TYPE   "
'
'            sStr = sStr & "    FROM ("
'            '---------------------------------------------------------------------------- 전체학생 조회 START
'            sStr = sStr & "          SELECT *"
'            sStr = sStr & "            FROM CLSTD01TB"
'            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'            sStr = sStr & "             AND EXMID > ' ' "
'            sStr = sStr & "             AND BIGO2 IS NULL "
'            sStr = sStr & "          UNION ALL"
'            '---------------------------------------------------------------------------- 전체학생 조회 END
'            '---------------------------------------------------------------------------- 합격자 조회 START
'            sStr = sStr & "          SELECT *"
'            sStr = sStr & "            From CLSTD01TB"
'            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
'            sStr = sStr & "             AND EXMID > ' ' "
'            sStr = sStr & "             AND BIGO2 IS NULL "
'            sStr = sStr & "          )"
'            sStr = sStr & "   GROUP BY SCHNO"
'            '---------------------------------------------------------------------------- 합격자 조회 END
'
'    sStr = sStr & "    ) "
'    sStr = sStr & " ORDER BY EXMID "
'
'
'    'Text1.Text = sStr
'    Set DBCmd = New ADODB.Command
'    Set DBRec = New ADODB.Recordset
'    Set DBParam = New ADODB.Parameter
'
'    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
'    DBCmd.CommandText = sStr
'    DBCmd.CommandType = adCmdText
'    DBCmd.CommandTimeout = 30
'
'
'
''>> 분원
''        sTmp = Trim(basModule.SchCD)
''        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'''>> 수험번호
''        If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''            sTmp = Trim(fpExmID_E.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
''            ' no action
''        End If
''>> 학생명
''        If Trim(txtStdNM.Text) > " " Then
''            sTmp = "%" & Trim(txtStdNM.Text) & "%"
''            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        End If
'
'
'    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
'    Do While DBRec.State And adStateExecuting
'        DoEvents
'    Loop
'
'    With DBRec
'        If .RecordCount = 0 Then
'
'            MsgBox "해당조회대상자가 없습니다.", vbExclamation + vbOKOnly, "전체학생 조회"
'
'        ElseIf .RecordCount > 0 Then
'
'            '## 헤더만들기
'            sprStdData.MaxRows = sprStdData.MaxRows + 1
'            sprStdData.Row = sprStdData.MaxRows
'
'            .MoveFirst
'            For ni = 0 To .Fields.count - 1 Step 1
'                sprStdData.Col = ni + 1
'                sTmp = " ":     If IsNull(.Fields(ni).Name) = False Then sTmp = Trim(.Fields(ni).Name)
'                    Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            Next ni
'
'            .MoveFirst
'            For nRec = 1 To .RecordCount Step 1
'                sprStdData.MaxRows = sprStdData.MaxRows + 1
'                sprStdData.Row = sprStdData.MaxRows
'
'
'                For ni = 0 To .Fields.count - 1 Step 1
'                    sprStdData.Col = ni + 1
'                    sTmp = " ":     If IsNull(.Fields(ni)) = False Then sTmp = Trim(.Fields(ni))
'                        Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
'                Next ni
'
'                .MoveNext
'
'            Next nRec
'
'
'        End If
'    End With
'
'    nRet = sprStdData.ExportToExcel(sExcelFileName, "Sheet1", sExcelLogFile)
'    MsgBox "엑셀자료 작성완료하였습니다.", vbInformation + vbOKOnly, "전체학생 조회"
'
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    Exit Sub
'
'ErrStmt1:
'    MsgBox "저장할 엑셀명을 등록하세요.", vbExclamation + vbOKOnly, Me.Caption
'    Exit Sub
'
'ErrStmt:
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    MsgBox "전체학생 조회시 에러가 발생하였습니다." & vbCrLf & _
'           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "전체학생 조회"
'
'    On Error GoTo 0
End Sub

















'학생 상세점수부분 등록 ==========================================================================

Private Sub Label56_Click()
    fraPoint.Visible = False
    
End Sub

Private Sub sprPoint_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        If sprPoint.MaxRows = 0 Then Exit Sub
        
        sprPoint.Row = sprPoint.ActiveRow
        
        sprPoint.DeleteRows sprPoint.Row, 1
        sprPoint.MaxRows = sprPoint.MaxRows - 1
    End If
    
End Sub


Private Sub sprPoint_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
    Dim sSubjCD     As String
    
    sprPoint.Row = Row
    sprPoint.Col = Col
    
    
    sSubjCD = Trim(Right(sprPoint.Text, 5))
    sprPoint.Col = 2
        Call basFunction.Set_SprType_Text(sprPoint, "center", "left", 10, sSubjCD)
        sprPoint.Lock = True
    
End Sub


Private Sub sprPoint_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    
    '데이터 조회
        
    Dim DBCmd           As ADODB.Command
    Dim DBRec           As ADODB.Recordset
    Dim DBParam         As ADODB.Parameter
    
    Dim nLength         As Long
    Dim sStr            As String
    Dim ni              As Integer
    
    Dim nRec            As Long
    
    
    Dim sTmp            As String
    Dim nTmp            As Long
    Dim nRet            As Long
    
    '> 초기화
    
    Dim sPyoJum         As String
    Dim sSubjCD         As String
    
    If cmdAddPoint.Tag = "ACC" Then Exit Sub
    If sprPoint.MaxRows = 0 Then Exit Sub
    
    sprPoint.Row = Row
    sprPoint.Col = 2
    
    If Trim(sprPoint.Text) = "" Then
        MsgBox "과목을 선택하세요.", vbExclamation + vbOKOnly, "학생표준점수 처리"
        Exit Sub
    Else
        sSubjCD = Trim(sprPoint.Text)
    End If
    
    sprPoint.Col = 4
        sPyoJum = Trim(sprPoint.Text)
        If Trim(sPyoJum) = "" Then sPyoJum = "0"
    
    On Error GoTo ErrStmt
    
    sStr = ""
    
    '김한욱 강남 요청 사항 강남의
    If Trim(basModule.schcd) = "K" Then
        sStr = sStr & " SELECT NVL(TO_CHAR(BAK_J),0) AS BAK_NUM"
    Else
        sStr = sStr & " SELECT NVL(TO_CHAR(DNG_J),0) AS BAK_NUM"
    End If
    
    sStr = sStr & "   FROM DMSITEMGR.DMINF28TB"
    sStr = sStr & "  WHERE YY = '2011'"
    sStr = sStr & "    AND TERM = '3'"
    sStr = sStr & "    AND CHA = '2'"
    sStr = sStr & "    AND GUBN = '1'"
    sStr = sStr & "    AND PRD_ID = 'U1011821'"
    sStr = sStr & "    AND PYO_J = '" & sPyoJum & "'"
    sStr = sStr & "    AND GM_CD = "
    sStr = sStr & "        ("
    sStr = sStr & "         SELECT GM_CD "
    sStr = sStr & "           FROM DMSITEMGR.DMINF19TB"
    sStr = sStr & "          WHERE DSHW_CD = '" & sSubjCD & "'"
    sStr = sStr & "         )"
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            
            .MoveFirst
            
            sprPoint.Col = 5
            
            sTmp = " "
            If IsNull(.Fields("BAK_NUM")) = False Then sTmp = Trim(.Fields("BAK_NUM"))
                Call basFunction.Set_SprType_Numeric(sprPoint, 0, 0, 99999, "", CDbl(sTmp))
                
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    sprPoint.Col = 5
            
    sTmp = "0"
    Call basFunction.Set_SprType_Text(sprPoint, "center", "left", basFunction.LenKor(sTmp), sTmp)
    
    On Error GoTo 0
    
    
    
End Sub













Private Sub cmdAddPointRow_Click()

    Dim sComboList      As String
    Dim sGbn            As String
    Dim sTmp            As String

    If Trim(txtSchNo.Text) = "" Then Exit Sub
    
    

    sprPoint.MaxRows = sprPoint.MaxRows + 1
    sprPoint.Row = sprPoint.MaxRows
    sprPoint.RowHeight(sprPoint.Row) = 15
    
    sprPoint.Col = 1
        sTmp = Trim(txtSchNo.Text)
            Call basFunction.Set_SprType_Text(sprPoint, "center", "left", basFunction.LenKor(sTmp), sTmp)
            sprPoint.Lock = True
            
    sprPoint.Col = 2
        sTmp = "37"
            Call basFunction.Set_SprType_Text(sprPoint, "center", "left", 10, sTmp)
            sprPoint.Lock = True
            
    sprPoint.Col = 3
        sprPoint.CellType = CellTypeComboBox
    
        sGbn = "CULT"
        
            Select Case Trim(basModule.schcd)
                Case "N", "B"
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "07", "09", "11", "13"
                            '<< 인문
                            sGbn = "CULT"
                            
                        Case "02", "04", "08", "10", "12", "14"
                            '<< 자연
                            sGbn = "SCI"
                            
                        Case "05", "15"
                            '<< 인문
                            sGbn = "CULT"
                        Case "06", "16"
                            '<< 자연
                            sGbn = "SCI"
                            
                        '2011-01-11 김한욱 노량진 부산에 의한 서울대,PS,PM 관련 추가
                        Case "21", "23"
                            '<< 인문
                            sGbn = "CULT"
                        Case "22", "24"
                            '<< 자연
                            sGbn = "SCI"
                    End Select
                
                    
                Case "S", "P", "J"
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "05", "11", "18"                                             '< 2008.02.15 : 계열 - 송파, 마송, 양재      2009.06.02 : 계열추가
                            '<< 인문
                            sGbn = "CULT"
                            
                        Case "02", "04", "06", "08", "12", "19"                                       '< 2008.02.15 : 계열 - 송파, 마송, 양재      2009.06.02 : 계열추가
                            '<< 자연
                            sGbn = "SCI"
                            
                    End Select
                Case Else
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "04", "06", "11", "16"                         '< 2008.01.10 : 계열 - 강남
                            '<< 인문
                            sGbn = "CULT"
                            
                        Case "02", "05", "07", "12", "17"                               '< 2008.01.10 : 계열 - 강남
                            '<< 자연
                            sGbn = "SCI"
                            
                    End Select
            End Select
    
            sComboList = ""
            
            If sGbn = "CULT" Then
                sComboList = sComboList & "언어                     37" + Chr$(9)
                sComboList = sComboList & "수리나형                 43" + Chr$(9)
                sComboList = sComboList & "외국어                   39" + Chr$(9)
                sComboList = sComboList & constSatams(0) & "                     " & constSatamCodes(0) + Chr$(9)
                sComboList = sComboList & constSatams(1) & "                     " & constSatamCodes(1) + Chr$(9)
                sComboList = sComboList & constSatams(2) & "                     " & constSatamCodes(2) + Chr$(9)
                sComboList = sComboList & constSatams(3) & "                     " & constSatamCodes(3) + Chr$(9)
                sComboList = sComboList & constSatams(4) & "                     " & constSatamCodes(4) + Chr$(9)
                sComboList = sComboList & constSatams(5) & "                     " & constSatamCodes(5) + Chr$(9)
                sComboList = sComboList & constSatams(6) & "                     " & constSatamCodes(6) + Chr$(9)
                sComboList = sComboList & constSatams(7) & "                     " & constSatamCodes(7) + Chr$(9)
                sComboList = sComboList & constSatams(8) & "                     " & constSatamCodes(8) + Chr$(9)
                sComboList = sComboList & constSatams(9) & "                     " & constSatamCodes(9) + Chr$(9)
                sComboList = sComboList & "독어                     31" + Chr$(9)
                sComboList = sComboList & "일어                     32" + Chr$(9)
                sComboList = sComboList & "에스파냐                 33" + Chr$(9)
                sComboList = sComboList & "불어                     34" + Chr$(9)
                sComboList = sComboList & "중국어                   35" + Chr$(9)
                sComboList = sComboList & "한문                     36" + Chr$(9)
                sComboList = sComboList & "아랍어                   42" + Chr$(9)
                sComboList = sComboList & "베트남어                 44"

            Else
                sComboList = sComboList & "언어                     37" + Chr$(9)
                sComboList = sComboList & "수리가형                 38" + Chr$(9)
                sComboList = sComboList & "외국어                   39" + Chr$(9)
                sComboList = sComboList & "물리1                    51" + Chr$(9)
                sComboList = sComboList & "화학1                    52" + Chr$(9)
                sComboList = sComboList & "생명과학1                    53" + Chr$(9)
                sComboList = sComboList & "지구과학1                54" + Chr$(9)
                sComboList = sComboList & "물리2                    55" + Chr$(9)
                sComboList = sComboList & "화학2                    56" + Chr$(9)
                sComboList = sComboList & "생명과학2                    57" + Chr$(9)
                sComboList = sComboList & "지구과학2                58" + Chr$(9)
                sComboList = sComboList & "미적분                   81" + Chr$(9)
                sComboList = sComboList & "이산수학                 82" + Chr$(9)
                sComboList = sComboList & "확률통계                 83"

            End If
    
        sprPoint.TypeComboBoxList = sComboList
        sprPoint.TypeComboBoxEditable = False
        sprPoint.TypeComboBoxMaxDrop = 11
        sprPoint.TypeComboBoxCurSel = 0
        sprPoint.TypeComboBoxWidth = 0
            
    sprPoint.Col = 4
        sTmp = "0"
            Call basFunction.Set_SprType_Numeric(sprPoint, 0, 0, 99999, "", CDbl(sTmp))
            
    sprPoint.Col = 5
        sTmp = "0"
            Call basFunction.Set_SprType_Numeric(sprPoint, 0, 0, 99999, "", CDbl(sTmp))
            
    sprPoint.Col = 6
        sprPoint.CellType = CellTypeButton
        sprPoint.TypeButtonText = "계산"
        
    sprPoint.Col = 7
        Call basFunction.Set_SprType_ChkBox(sprPoint)
        sprPoint.value = 0
        
    sprPoint.Col = 8
        Call basFunction.Set_SprType_ChkBox(sprPoint)
        sprPoint.value = 1
        
    sprPoint.Col = 9
        Call basFunction.Set_SprType_ChkBox(sprPoint)
        sprPoint.value = 0
        
End Sub


Private Sub cmdSavePoint_Click()
    
    If sprPoint.MaxRows = 0 Then
        MsgBox "등록할 내용이 없습니다.", vbExclamation + vbOKOnly, "학생점수등록"
        Exit Sub
    End If
    
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    Dim sStr        As String
    
    Dim sSchNO      As String
    Dim sSubID      As String
    Dim sSubNum     As String
    Dim sSubBak     As String
    
    Dim nRow        As Long
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nExe        As Integer
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    
    nExe = 0
    
    For nRow = 1 To sprPoint.MaxRows Step 1
    
        sprPoint.Row = nRow
        
        sStr = ""
        
        sprPoint.Col = 1:       sSchNO = Trim(sprPoint.Text)
        sprPoint.Col = 2:       sSubID = Trim(sprPoint.Text)
        sprPoint.Col = 4:       sSubNum = Trim(sprPoint.Text)
        sprPoint.Col = 5:       sSubBak = Trim(sprPoint.Text)
        
            sprPoint.Col = 7                '삭제처리
            If sprPoint.value = 1 Then
                
                sStr = sStr & " DELETE CLSTD03TB "
                sStr = sStr & "  WHERE SCHNO   = '" & sSchNO & "'"
                sStr = sStr & "    AND SUB_ID  = '" & sSubID & "'"
            Else
                
                sprPoint.Col = 9
                If sprPoint.value = 1 Then              '갱신등록
                    
                    sStr = sStr & " UPDATE CLSTD03TB "
                    sStr = sStr & "    SET SUB_NUM = '" & sSubNum & "', "
                    sStr = sStr & "        SUB_BAK = '" & sSubBak & "' "
                    sStr = sStr & "  WHERE SCHNO   = '" & sSchNO & "'"
                    sStr = sStr & "    AND SUB_ID  = '" & sSubID & "'"
                    
                Else
                
                    sprPoint.Col = 8                    '신규등록
                    If sprPoint.value = 1 Then
                    
                        sStr = sStr & " INSERT INTO CLSTD03TB (SCHNO, SUB_ID, SUB_NUM, SUB_BAK) "
                        sStr = sStr & " VALUES ( "
                        sStr = sStr & "         '" & sSchNO & "',"
                        sStr = sStr & "         '" & sSubID & "',"
                        sStr = sStr & "         '" & sSubNum & "',"
                        sStr = sStr & "         '" & sSubBak & "' "
                        sStr = sStr & "        )"
                        
                    End If
                End If
                
            End If
        
        If sStr > " " Then
        
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
            
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
        End If
        
    Next nRow
    
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    MsgBox "등록하였습니다.", vbInformation + vbOKOnly, "학생점수 등록"
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생점수 등록"
    
End Sub













Private Sub cmdAddPoint_Click()
    
    '데이터 조회
        
    Dim DBCmd           As ADODB.Command
    Dim DBRec           As ADODB.Recordset
    Dim DBParam         As ADODB.Parameter
    
    Dim nLength         As Long
    Dim sStr            As String
    Dim ni              As Integer
    
    Dim nRec            As Long
    
    
    Dim sTmp            As String
    Dim nTmp            As Long
    Dim nRet            As Long
    
    cmdAddPoint.Tag = "ACC"
    
    
    '> 초기화
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "학생을 선택하세요.", vbExclamation + vbOKOnly, "학생 상세점수"
        Exit Sub
    End If
    
    sprPoint.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, SUB_ID,"
    
    sStr = sStr & "         CASE WHEN      INSTR(SUB_ID,'37') > 0 THEN     /* 언어 */"
    sStr = sStr & "             '언어'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'38') > 0 THEN     /* 수리가형 */"
    sStr = sStr & "             '수리가형'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'39') > 0 THEN     /* 외국어 */"
    sStr = sStr & "             '외국어'"
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(0) & "') > 0 THEN     /* 사탐-한국사 */"
    sStr = sStr & "             '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'51') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "             '물리1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(1) & "') > 0 THEN     /* 사탐-윤리 */"
    sStr = sStr & "             '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'52') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "             '화학1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(2) & "') > 0 THEN     /* 사탐-경제 */"
    sStr = sStr & "             '" & constSatams(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'53') > 0 THEN     /* 과탐-생명과학1 */"
    sStr = sStr & "             '생명과학1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(3) & "') > 0 THEN     /* 사탐-한국근현대 */"
    sStr = sStr & "             '" & constSatams(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'54') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "             '지학1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(4) & "') > 0 THEN     /* 사탐-세계사 */"
    sStr = sStr & "             '" & constSatams(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'55') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "             '물리2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(5) & "') > 0 THEN     /* 사탐-경제지리 */"
    sStr = sStr & "             '" & constSatams(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'56') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "             '화학2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(6) & "') > 0 THEN     /* 사탐-한국지리 */"
    sStr = sStr & "             '" & constSatams(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'57') > 0 THEN     /* 과탐-생명과학2 */"
    sStr = sStr & "             '생명과학2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(7) & "') > 0 THEN     /* 사탐-정치 */"
    sStr = sStr & "             '" & constSatams(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'58') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "             '지학2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(8) & "') > 0 THEN     /* 사탐-사회문화 */"
    sStr = sStr & "             '" & constSatams(8) & "'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(9) & "') > 0 THEN     /* 사탐-법과사회 */"
    sStr = sStr & "             '" & constSatams(9) & "'"
    sStr = sStr & "      /* 제2외국어 & 수리 */"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'31') > 0 THEN '독어'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'32') > 0 THEN '일어'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'33') > 0 THEN '에스파냐'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'34') > 0 THEN '불어'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'35') > 0 THEN '중국어'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'36') > 0 THEN '한문'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'42') > 0 THEN '아랍어'"
                                                                                                                 
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'81') > 0 THEN '미적분'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'82') > 0 THEN '이산수학'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'83') > 0 THEN '확률통계'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'43') > 0 THEN '수리나형'"
    
    sStr = sStr & "         END END END END END END END END END END END END END END END END END END END END END END END END END END END END END END END END SUBJNM, "
    
    sStr = sStr & "         SUB_NUM, SUB_BAK"
    sStr = sStr & "    From CLSTD03TB"
    sStr = sStr & "   WHERE SCHNO = '" & Trim(txtSchNo) & "'"
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            sprPoint.MaxRows = .RecordCount
            
            .MoveFirst
            For nRec = 1 To .RecordCount Step 1
                sprPoint.Row = nRec
                sprPoint.RowHeight(nRec) = 15
                
                sprPoint.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprPoint, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        
                sprPoint.Col = 2
                    sTmp = " ":     If IsNull(.Fields("SUB_ID")) = False Then sTmp = Trim(.Fields("SUB_ID"))
                        Call basFunction.Set_SprType_Text(sprPoint, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        
                sprPoint.Col = 3
                    sTmp = " ":     If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprPoint, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        
                sprPoint.Col = 4
                    sTmp = "0":     If IsNull(.Fields("SUB_NUM")) = False Then sTmp = Trim(.Fields("SUB_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprPoint, 0, 0, 99999, "", CDbl(sTmp))
                        
                sprPoint.Col = 5
                    sTmp = "0":     If IsNull(.Fields("SUB_BAK")) = False Then sTmp = Trim(.Fields("SUB_BAK"))
                        If sTmp <> "X" Then
                            Call basFunction.Set_SprType_Numeric(sprPoint, 0, 0, 99999, "", CDbl(sTmp))
                            
                        End If
                        
                sprPoint.Col = 6
                    sprPoint.CellType = CellTypeButton
                    sprPoint.TypeButtonText = "계산"
                    
                sprPoint.Col = 7
                    Call basFunction.Set_SprType_ChkBox(sprPoint)
                    sprPoint.value = 0
                    
                sprPoint.Col = 8
                    Call basFunction.Set_SprType_ChkBox(sprPoint)
                    sprPoint.value = 0
                    
                sprPoint.Col = 9
                    Call basFunction.Set_SprType_ChkBox(sprPoint)
                    sprPoint.value = 1
                    
                
                .MoveNext
                
            Next nRec
            
                    
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdAddPoint.Tag = ""
    fraPoint.Visible = True
    
    Exit Sub
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "학생자료 조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "학생점수조회"
    
    On Error GoTo 0
    
    fraPoint.Visible = True
    
    cmdAddPoint.Tag = ""
    
End Sub














'## 상세항목 조회
Private Sub Label49_Click()         '< 닫기
    fraAddr.Visible = False
    
End Sub

Private Sub cmdChgAddr_Click()
    fraAddr.Visible = True
    fpBirth_ymdS.SetFocus
    
End Sub



Private Sub cmdSaveAddr_Click()
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim ni          As Integer
    Dim nExe        As Long
    
    Dim sBirth_ymd      As String
    Dim sZip        As String
    Dim sAddr1      As String
    Dim sAddr2      As String
    Dim sEmail      As String
    
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "학생을 조회하세요.", vbExclamation + vbOKOnly, "상세내역 변경"
        Exit Sub
    End If
    
    sBirth_ymd = Trim(fpBirth_ymdS.UnFmtText)
    sZip = Trim(fpZip.UnFmtText)
    sAddr1 = Trim(txtAddr1.Text)
    sAddr2 = Trim(txtAddr2.Text)
    sEmail = Trim(txtEmail.Text)
    
    If MsgBox("【 " & Trim(txtStdNM.Text) & " 】" & " 학생의 상세내역을 변경하시겠습니까?", vbQuestion + vbYesNo, "상세내역 변경") = vbNo Then
        Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    
    
    sStr = ""
    sStr = sStr & "  UPDATE CLSTD01TB"
    sStr = sStr & "     SET Birth_ymd = '" & Trim(fpBirth_ymdS.UnFmtText) & "',"
    sStr = sStr & "         ZIP   = '" & Trim(fpZip.UnFmtText) & "',"
    sStr = sStr & "         ADDR1 = '" & Trim(txtAddr1.Text) & "',"
    sStr = sStr & "         ADDR2 = '" & Trim(txtAddr2.Text) & "',"
    sStr = sStr & "         EMAIL = '" & Trim(txtEmail.Text) & "'"
    sStr = sStr & "   WHERE SCHNO = '" & Trim(txtSchNo.Text) & "'"
    
    nExe = 0
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "등록하였습니다.", vbInformation + vbOKOnly, "상세내역 변경"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "상세내역 변경"
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "상세내역 변경"
End Sub















'## 선택항목만 받기
Private Sub imgExcel_Click()
    
    Dim DBCmd           As ADODB.Command
    Dim DBRec           As ADODB.Recordset
    Dim DBParam         As ADODB.Parameter
    
    Dim nLength         As Long
    Dim sStr            As String
    Dim ni              As Integer
    
    Dim nRec            As Long
    
    
    Dim sTmp            As String
    Dim nTmp            As Long
    Dim nRet            As Long
    
    Dim sExcelFileName  As String
    Dim sExcelLogFile   As String
    
    '> 초기화
    sprStdData.MaxRows = 0
    
    On Error GoTo ErrStmt1
    
    With dlgFile
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path
        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowSave
        
        If (.fileName) = "" Then
            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        sExcelFileName = .fileName
        
        ni = InStrRev(sExcelFileName, "\", -1, vbTextCompare)
        sExcelLogFile = Mid(sExcelFileName, 1, ni) & "\" & Mid(sExcelFileName, ni + 1, Len(sExcelFileName) - ni + 1 - 5)
        
    End With
    
    On Error GoTo 0
    
    On Error GoTo ErrStmt
    
    '학생 엑셀 저장 쿼리문
    sStr = basCommonSTD.Get_StdExcuteSqlToExcel_N(cboKaeyol_F.Text)
    
    
    Text1.Text = sStr
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
            
            MsgBox "해당조회대상자가 없습니다.", vbExclamation + vbOKOnly, "전체학생 조회"
            
        ElseIf .RecordCount > 0 Then
            
            '## 헤더만들기
            sprStdData.MaxRows = sprStdData.MaxRows + 1
            sprStdData.Row = sprStdData.MaxRows
                
            .MoveFirst
            For ni = 0 To .Fields.count - 1 Step 1
                sprStdData.Col = ni + 1
                sTmp = " ":     If IsNull(.Fields(ni).Name) = False Then sTmp = Trim(.Fields(ni).Name)
                    Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
            Next ni
            
            .MoveFirst
            For nRec = 1 To .RecordCount Step 1
                sprStdData.MaxRows = sprStdData.MaxRows + 1
                sprStdData.Row = sprStdData.MaxRows
                
                
                For ni = 0 To .Fields.count - 1 Step 1
                    sprStdData.Col = ni + 1
                    sTmp = " ":     If IsNull(.Fields(ni)) = False Then sTmp = Trim(.Fields(ni))
                        Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                Next ni
                
                .MoveNext
                
            Next nRec
            
                    
        End If
    End With
    
    nRet = sprStdData.ExportToExcel(sExcelFileName, "Sheet1", sExcelLogFile)
    MsgBox "엑셀자료 작성완료하였습니다.", vbInformation + vbOKOnly, "전체학생 조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
    
ErrStmt1:
    MsgBox "저장할 엑셀명을 등록하세요.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "전체학생 조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "전체학생 조회"
    
    On Error GoTo 0
End Sub


'
''## 선택항목만 받기
'Private Sub imgExcel_Click()
'
'    Dim DBCmd           As ADODB.Command
'    Dim DBRec           As ADODB.Recordset
'    Dim DBParam         As ADODB.Parameter
'
'    Dim nLength         As Long
'    Dim sStr            As String
'    Dim ni              As Integer
'
'    Dim nRec            As Long
'
'
'    Dim sTmp            As String
'    Dim nTmp            As Long
'    Dim nRet            As Long
'
'    Dim sExcelFileName  As String
'    Dim sExcelLogFile   As String
'
'    '> 초기화
'    sprStdData.MaxRows = 0
'
'    On Error GoTo ErrStmt1
'
'    With dlgFile
'        .CancelError = True
'        .fileName = ""
'        .InitDir = App.Path
'        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
'        .DefaultExt = "*.XLS"
'        .ShowSave
'
'        If (.fileName) = "" Then
'            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
'            Exit Sub
'        End If
'
'        sExcelFileName = .fileName
'
'        ni = InStrRev(sExcelFileName, "\", -1, vbTextCompare)
'        sExcelLogFile = Mid(sExcelFileName, 1, ni) & "\" & Mid(sExcelFileName, ni + 1, Len(sExcelFileName) - ni + 1 - 5)
'
'    End With
'
'    On Error GoTo 0
'
'    On Error GoTo ErrStmt
'
'    sStr = ""
'    sStr = sStr & "  SELECT SCHNO AS 시스템코드   , "
'    sStr = sStr & "         ACID  AS 학원   , "
'    sStr = sStr & "         EXMID AS 수험번호, STDNM AS 학생, D_UNIVCD AS 지원대학, D_MAJORCD AS 지원단대,"
'
'    If basModule.SchCD = "N" Then
'        Select Case basModule.RegID
'            Case "10000", "00002", "10003", "00001" '김영덕과장
'                sStr = sStr & "         SUBSTR(REPLACE(Birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(Birth_ymd,'-',''),7,7) AS 주민번호,"
'            Case "10001"                            '신현우
'                sStr = sStr & "         SUBSTR(REPLACE(Birth_ymd,'-',''),1,6)||'-*******' AS 주민번호,"
'            Case "10002"                            '정순택
'                sStr = sStr & "         SUBSTR(REPLACE(Birth_ymd,'-',''),1,6)||'-*******' AS 주민번호,"
'        End Select
'    Else
'        sStr = sStr & "         SUBSTR(REPLACE(Birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(Birth_ymd,'-',''),7,7) AS 주민번호,"
'    End If
'
'    sStr = sStr & "         DECODE(EXMTYPE,'0','무시험','1','유시험') AS 시험형태, "
'    sStr = sStr & "         DECODE(KAEYOL,'01','인문',"
'    sStr = sStr & "                       '02','자연',"
''<< 계열 >> : 2008.01.09
'    If Trim(basModule.SchCD) = "N" Then
'        sStr = sStr & "                   '03','예체',"
'        sStr = sStr & "                   '04','수리(나)',"
'        sStr = sStr & "                   '05','인문수능',"
'        sStr = sStr & "                   '06','자연수능',"
'
'        sStr = sStr & "                   '07','신설인문',"
'        sStr = sStr & "                   '08','신설자연',"
'        sStr = sStr & "                   '09','신설수능인문',"
'        sStr = sStr & "                   '10','신설수능자연',"
'
'        sStr = sStr & "                   '11','편)인문',"
'        sStr = sStr & "                   '12','편)자연',"
'        sStr = sStr & "                   '13','편)예체',"
'        sStr = sStr & "                   '14','편)수리(나)',"
'        sStr = sStr & "                   '15','편)인문수능',"
'        sStr = sStr & "                   '16','편)자연수능',"
'        sStr = sStr & "                   '21','서울대인문',"
'        sStr = sStr & "                   '22','서울대자연',"
'    End If
''<< 계열 >> : 2008.01.10
'    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
'        sStr = sStr & "                   '04','주말법대',"
'        sStr = sStr & "                   '05','주말의대',"
'        sStr = sStr & "                   '06','야간법대',"
'        sStr = sStr & "                   '07','야간의대',"
'
'        sStr = sStr & "                   '11','선착순인문',"
'        sStr = sStr & "                   '12','선착순자연',"
'
'        sStr = sStr & "                   '16','선착순인문16',"
'        sStr = sStr & "                   '17','선착순자연17',"
'    End If
''<< 계열 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "S" Then
'        sStr = sStr & "                   '03','예체능',"
'        'sStr = sStr & "                   '04','특별자연',"
'
'        sStr = sStr & "                   '05','수능인문',"
'        sStr = sStr & "                   '06','수능자연',"
'
'        sStr = sStr & "                   '11','신설인문',"
'        sStr = sStr & "                   '12','신설자연',"
'
'        sStr = sStr & "                   '18','인문프리미엄',"
'        sStr = sStr & "                   '19','자연프리미엄',"
'        sStr = sStr & "                   '21','서울대특별인문',"
'        sStr = sStr & "                   '22','서울대특별자연',"
'        sStr = sStr & "                   '23','야간서울대인문',"
'        sStr = sStr & "                   '24','야간서울대자연',"
'
'    End If
''<< 계열 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "P" Then         '< 마송
'        sStr = sStr & "                   '03','특별인문',"
'        sStr = sStr & "                   '04','특별자연',"
'    End If
'
'    If Trim(basModule.SchCD) = "J" Then         '< 양재
'        sStr = sStr & "                   '11','신설인문',"
'        sStr = sStr & "                   '12','신설자연',"
'
'        sStr = sStr & "                   '18','인문프리미엄',"
'        sStr = sStr & "                   '19','자연프리미엄',"
'    End If
'
''<< 계열 >> : 2009.01.09
'    If Trim(basModule.SchCD) = "B" Then         '< 부산
'        sStr = sStr & "                   '05','특별인문',"
'        sStr = sStr & "                   '06','특별자연',"
'        sStr = sStr & "                   '07','연고대인문',"
'        sStr = sStr & "                   '08','연고대자연',"
'        sStr = sStr & "                   '09','심화인문',"
'        sStr = sStr & "                   '10','심화자연',"
'    End If
'
'    sStr = sStr & "                       '','기타') AS 계열,"
'
'    sStr = sStr & "     /* 사탐, 과탐 분리 */"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* 사탐-국사 */"
'    sStr = sStr & "             '국사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
'    sStr = sStr & "             '물1'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구1,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* 사탐-윤리 */"
'    sStr = sStr & "             '윤리'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
'    sStr = sStr & "             '화1'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구2,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* 사탐-경제 */"
'    sStr = sStr & "             '경제'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생명과학1 */"
'    sStr = sStr & "             '생1'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구3,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* 사탐-한국근현대 */"
'    sStr = sStr & "             '한근'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
'    sStr = sStr & "             '지1'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구4,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* 사탐-세계사 */"
'    sStr = sStr & "             '세사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
'    sStr = sStr & "             '물2'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구5,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* 사탐-경제지리 */"
'    sStr = sStr & "             '경지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
'    sStr = sStr & "             '화2'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구6,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* 사탐-한국지리 */"
'    sStr = sStr & "             '한지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생명과학2 */"
'    sStr = sStr & "             '생2'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구7,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* 사탐-정치 */"
'    sStr = sStr & "             '정치'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
'    sStr = sStr & "             '지2'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END AS 탐구8,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* 사탐-사회문화 */"
'    sStr = sStr & "             '사문'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END AS 탐구9,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* 사탐-법과사회 */"
'    sStr = sStr & "             '법사'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END AS 탐구10,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* 사탐-세계지리 */"
'    sStr = sStr & "             '세지'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END AS 탐구11,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* 제2외국어 & 수리 */"
'    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에파'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
'
'    '<< 송파 >> : 2008.01.09
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '언어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '수리'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '영어'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '세계사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '세지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '아랍어'"
'
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '나형'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END END END END END END END END END END END END END END END 제2선택,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* 논술 */"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
'    sStr = sStr & "             '언어'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 언어논술,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
'    sStr = sStr & "             '수리'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 수리논술,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
'    sStr = sStr & "             '외국어'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 사탐논술,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< 변경
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END 과탐논술,"
'    sStr = sStr & "  "
'    sStr = sStr & "         CY_ACNT AS 가상계좌, TOT_AMT AS 전체금액    ,"
'    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS 기본금액1  ,"
'    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS 기본금액2  ,"
'    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS 기본금액3  ,"
'    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS 기본금액4  ,"
'    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS 기본금액5  ,"
'    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS 기본금액6  ,"
'    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS 기본금액7  ,"
'    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS 기본금액8  ,"
'    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS 탐구영역금액1 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS 탐구영역금액2 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS 탐구영역금액3 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS 탐구영역금액4 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS 탐구영역금액5 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS 탐구영역금액6 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS 탐구영역금액7 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS 탐구영역금액8 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS 탐구영역금액9 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS 탐구영역금액10,"
'    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS 탐구영역금액11,"
'    sStr = sStr & "      /* 탐구 성적 떄문에 처리.. */"
'    sStr = sStr & "              CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'01|') > 0 THEN '국사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'02|') > 0 THEN '윤리'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'03|') > 0 THEN '경제'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'04|') > 0 THEN '한근'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'05|') > 0 THEN '세계사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'06|') > 0 THEN '경지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'07|') > 0 THEN '한지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'08|') > 0 THEN '정치'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'09|') > 0 THEN '사문'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'10|') > 0 THEN '법사'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'11|') > 0 THEN '세지'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'51|') > 0 THEN '물I'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'52|') > 0 THEN '화I'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'53|') > 0 THEN '생I'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'54|') > 0 THEN '지I'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'55|') > 0 THEN '물II'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'56|') > 0 THEN '화II'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'57|') > 0 THEN '생II'"
'    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'58|') > 0 THEN '지II'"
'    sStr = sStr & "         END END END END END END END END END END END END END END END END END END SEL_X6,"
'    sStr = sStr & "         K_NUM AS 언어점수, M_NUM AS 수학점수, E_NUM AS 영어점수, "
'    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS 전체점수,"
'
'
'    sStr = sStr & "         DECODE(SEL1_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제1지망,"
'    sStr = sStr & "         DECODE(SEL2_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제2지망,"
'
'    sStr = sStr & "         DECODE(PASS1,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격1   ,"
'    sStr = sStr & "         DECODE(PASS2,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격2   ,"
'    sStr = sStr & "         DECODE(PASS3,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격3   ,"
'    sStr = sStr & "         DECODE(PASS4,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격4   ,"
'
'
'    sStr = sStr & "         DECODE(SEX,'M','남','F','여') AS 성별        , "
'    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS 우편번호, ADDR1 AS 우편주소      , ADDR2 AS 상세주소     ,"
'    sStr = sStr & "         TEL AS 전화번호, CEL AS 핸드폰        , EMAIL AS 이메일     ,"
'    sStr = sStr & "         HIGH_SCH AS 고등학교 , GRADE_YEAR AS 졸업년도 ,"
'    sStr = sStr & "         PRNT_NM AS 학부모명 , DECODE(PRNT_RLTN,'1','부','2','모','3','기타') AS 학부모관계, "
'    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS 학부모_우편번호, PRNT_ADDR1 AS 학부모_우편주소 , PRNT_ADDR2 AS 학부모_상세주소,"
'    sStr = sStr & "         PRNT_TEL AS 학부모_전화번호  , PRNT_CEL AS 학부모_핸드폰   , PRNT_JOB AS 학부모_직업   , PRNT_W_TEL AS 학부모_직장전화 ,"
'    sStr = sStr & "         PHOTO_PATH AS 사진저장장소, "
'    sStr = sStr & "         DECODE(R_WAY,'1','학원등록','2','인터넷등록','3','학원등록') AS 등록번호, "
'    sStr = sStr & "         ORD_NO AS 주문번호, "
'    sStr = sStr & "         ACID||EXMID AS 이미지파일명, "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
'    sStr = sStr & "         REGDATE AS 등록일자, GET_PAYGUBN(ORD_NO) AS 결재방법, CASH_BILL_NUM AS 현금영수증,"
'    sStr = sStr & "         DECODE(MU_TYPE,'1','수능평가','2','6월 평가원','3','9월 평가원','4','6월 평가원','5','9월 평가원','') AS 등급, "
'    sStr = sStr & "         CL_CLOSE AS 완료년월 "
'
'    sStr = sStr & " , "
'        sStr = sStr & "        J01 AS 언어          ,"
'        sStr = sStr & "        K01 AS 언어_백       ,"
'        sStr = sStr & "        J02 AS 수리가        ,"
'        sStr = sStr & "        K02 AS 수리가형_백   ,"
'        sStr = sStr & "        J03 AS 외국어        ,"
'        sStr = sStr & "        K03 AS 외국어_백     ,"
'
'        sStr = sStr & "        J04 AS 국사_물1      ,"
'        sStr = sStr & "        K04 AS 국사_물1_백   ,"
'        sStr = sStr & "        J05 AS 윤리_화1      ,"
'        sStr = sStr & "        K05 AS 윤리_화1_백   ,"
'        sStr = sStr & "        J06 AS 경제_생1      ,"
'        sStr = sStr & "        K06 AS 경제_생1_백   ,"
'        sStr = sStr & "        J07 AS 한근_지학1    ,"
'        sStr = sStr & "        K07 AS 한근_지학1_백 ,"
'        sStr = sStr & "        J08 AS 세사_물2      ,"
'        sStr = sStr & "        K08 AS 세사_물2_백   ,"
'        sStr = sStr & "        J09 AS 경지_화2      ,"
'        sStr = sStr & "        K09 AS 경지_화2_백   ,"
'        sStr = sStr & "        J10 AS 한지_생2      ,"
'        sStr = sStr & "        K10 AS 한지_생2_백   ,"
'        sStr = sStr & "        J11 AS 정치_지학2    ,"
'        sStr = sStr & "        K11 AS 정치_지학2_백 ,"
'
'        sStr = sStr & "        J12 AS 사문          ,"
'        sStr = sStr & "        K12 AS 사문_백       ,"
'        sStr = sStr & "        J13 AS 법사          ,"
'        sStr = sStr & "        K13 AS 법사_백       ,"
'        sStr = sStr & "        J14 AS 세지          ,"
'        sStr = sStr & "        K14 AS 세지_백       ,"
'
'        sStr = sStr & "        J15 AS 독어_미적     ,"
'        sStr = sStr & "        K15 AS 독어_미적_백  ,"
'        sStr = sStr & "        J16 AS 일어_이산     ,"
'        sStr = sStr & "        K16 AS 일어_이산_백  ,"
'        sStr = sStr & "        J17 AS 에파_확통     ,"
'        sStr = sStr & "        K17 AS 에파_확통_백  ,"
'        sStr = sStr & "        J18 AS 불어_수리나   ,"
'        sStr = sStr & "        K18 AS 불어_수리나_백,"
'
'        sStr = sStr & "        J19 AS 중어          ,"
'        sStr = sStr & "        K19 AS 중어_백       ,"
'        sStr = sStr & "        J20 AS 한문          ,"
'        sStr = sStr & "        K20 AS 한문_백       ,"
'        sStr = sStr & "        J21 AS 아랍어        ,"
'        sStr = sStr & "        K21 AS 아랍어_백     "
'
'    sStr = sStr & "    FROM ( "
'
'            sStr = sStr & "  SELECT A.SCHNO           ,"
'            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
'            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
'            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
'            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
'            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
'            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
'            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
'            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      , MAX(SEL6      ) AS  SEL6      ,"
'            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   ,"
'            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
'            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
'            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
'            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
'            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
'            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
'            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
'            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
'            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
'            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
'            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
'            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
'            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
'            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
'            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
'            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(MU_TYPE   ) AS MU_TYPE   , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM"
'
'                    sStr = sStr & " , "
'                    sStr = sStr & "        SUM(J01) AS J01,"
'                    sStr = sStr & "        SUM(K01) AS K01,"
'                    sStr = sStr & "        SUM(J02) AS J02,"
'                    sStr = sStr & "        SUM(K02) AS K02,"
'                    sStr = sStr & "        SUM(J03) AS J03,"
'                    sStr = sStr & "        SUM(K03) AS K03,"
'
'                    sStr = sStr & "        SUM(J04) AS J04,"
'                    sStr = sStr & "        SUM(K04) AS K04,"
'                    sStr = sStr & "        SUM(J05) AS J05,"
'                    sStr = sStr & "        SUM(K05) AS K05,"
'                    sStr = sStr & "        SUM(J06) AS J06,"
'                    sStr = sStr & "        SUM(K06) AS K06,"
'                    sStr = sStr & "        SUM(J07) AS J07,"
'                    sStr = sStr & "        SUM(K07) AS K07,"
'                    sStr = sStr & "        SUM(J08) AS J08,"
'                    sStr = sStr & "        SUM(K08) AS K08,"
'                    sStr = sStr & "        SUM(J09) AS J09,"
'                    sStr = sStr & "        SUM(K09) AS K09,"
'                    sStr = sStr & "        SUM(J10) AS J10,"
'                    sStr = sStr & "        SUM(K10) AS K10,"
'                    sStr = sStr & "        SUM(J11) AS J11,"
'                    sStr = sStr & "        SUM(K11) AS K11,"
'
'                    sStr = sStr & "        SUM(J12) AS J12,"
'                    sStr = sStr & "        SUM(K12) AS K12,"
'                    sStr = sStr & "        SUM(J13) AS J13,"
'                    sStr = sStr & "        SUM(K13) AS K13,"
'                    sStr = sStr & "        SUM(J14) AS J14,"
'                    sStr = sStr & "        SUM(K14) AS K14,"
'
'                    sStr = sStr & "        SUM(J15) AS J15,"
'                    sStr = sStr & "        SUM(K15) AS K15,"
'                    sStr = sStr & "        SUM(J16) AS J16,"
'                    sStr = sStr & "        SUM(K16) AS K16,"
'                    sStr = sStr & "        SUM(J17) AS J17,"
'                    sStr = sStr & "        SUM(K17) AS K17,"
'                    sStr = sStr & "        SUM(J18) AS J18,"
'                    sStr = sStr & "        SUM(K18) AS K18,"
'
'                    sStr = sStr & "        SUM(J19) AS J19,"
'                    sStr = sStr & "        SUM(K19) AS K19,"
'                    sStr = sStr & "        SUM(J20) AS J20,"
'                    sStr = sStr & "        SUM(K20) AS K20,"
'                    sStr = sStr & "        SUM(J21) AS J21,"
'                    sStr = sStr & "        SUM(K21) AS K21"
'
'            sStr = sStr & "    FROM ("
'            '---------------------------------------------------------------------------- 전체학생 조회 START
'            sStr = sStr & "          SELECT *"
'            sStr = sStr & "            FROM CLSTD01TB"
'            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'            sStr = sStr & "             AND EXMID > ' ' "
'
'    If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then
'            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
'    End If
'            sStr = sStr & "             AND BIGO2 IS NULL "
'            sStr = sStr & "          UNION ALL"
'            '---------------------------------------------------------------------------- 전체학생 조회 END
'            '---------------------------------------------------------------------------- 합격자 조회 START
'            sStr = sStr & "          SELECT *"
'            sStr = sStr & "            From CLSTD01TB"
'            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
'            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
'            sStr = sStr & "             AND EXMID > ' ' "
'    If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then
'            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
'    End If
'            sStr = sStr & "             AND BIGO2 IS NULL "
'
'            sStr = sStr & "          ) A, "
'
'            sStr = sStr & "               ("
'
'                sStr = sStr & "         SELECT SCHNO,"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* 언어                  */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* 백분위  언어          */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* 수리가형              */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* 백분위  수리가형      */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* 외국어                */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* 백분위  외국어        */"
'
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '01', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* 사탐-국사        , 과탐-물리1             */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '01', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* 백분위  사탐-국사        , 과탐-물리1     */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '02', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* 사탐-윤리        , 과탐-화학1             */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '02', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* 백분위  사탐-윤리        , 과탐-화학1     */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '03', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* 사탐-경제        , 과탐-생명과학1             */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '03', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* 백분위  사탐-경제        , 과탐-생명과학1     */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '04', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* 사탐-한국근현대  , 과탐-지구과학1         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '04', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* 백분위  사탐-한국근현대  , 과탐-지구과학1 */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '05', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* 사탐-세계사      , 과탐-물리2             */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '05', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* 백분위  사탐-세계사      , 과탐-물리2     */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '06', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* 사탐-경제지리    , 과탐-화학2             */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '06', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* 백분위  사탐-경제지리    , 과탐-화학2     */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '07', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* 사탐-한국지리    , 과탐-생명과학2           */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '07', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* 백분위 사탐-한국지리    , 과탐-생명과학2    */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '08', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* 사탐-정치        , 과탐-지구과학2         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '08', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* 백분위  사탐-정치        , 과탐-지구과학2 */"
'
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '09', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* 사탐-사회문화         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '09', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* 백분위  사탐-사회문화 */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '10', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* 사탐-법과사회         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '10', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* 백분위  사탐-법과사회 */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '11', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J14,    /* 사탐-세계지리         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '11', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K14,    /* 백분위  사탐-세계지리 */"
'
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* 독어             , 미적분                 */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* 백분위  독어             , 미적분         */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* 일어             , 이산수학               */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* 백분위  일어             , 이산수학       */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* 에스파냐         , 확률통계               */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* 백분위  에스파냐         , 확률통계       */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* 불어             , 수리나형               */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* 백분위  불어             , 수리나형       */"
'
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* 중국어                */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* 백분위  중국어        */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* 한문                  */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* 백분위  한문          */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* 아랍어                */"
'                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* 백분위  아랍어        */"
'                sStr = sStr & "           FROM CLSTD03TB"
'
'        sStr = sStr & "                ) B"
'        sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
'
'            sStr = sStr & "   GROUP BY A.SCHNO"
'            '---------------------------------------------------------------------------- 합격자 조회 END
'
'    sStr = sStr & "    ) "
'    sStr = sStr & " ORDER BY EXMID "
'
'
'
'    Set DBCmd = New ADODB.Command
'    Set DBRec = New ADODB.Recordset
'    Set DBParam = New ADODB.Parameter
'
'    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
'    DBCmd.CommandText = sStr
'    DBCmd.CommandType = adCmdText
'    DBCmd.CommandTimeout = 30
'


'
''>> 분원
''        sTmp = Trim(basModule.SchCD)
''        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'''>> 수험번호
''        If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''            sTmp = Trim(fpExmID_E.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
''            sTmp = Trim(fpExmID_S.UnFmtText)
''                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
''            ' no action
''        End If
''>> 학생명
''        If Trim(txtStdNM.Text) > " " Then
''            sTmp = "%" & Trim(txtStdNM.Text) & "%"
''            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        End If
'
'
'    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
'    Do While DBRec.State And adStateExecuting
'        DoEvents
'    Loop
'
'    With DBRec
'        If .RecordCount = 0 Then
'
'            MsgBox "해당조회대상자가 없습니다.", vbExclamation + vbOKOnly, "전체학생 조회"
'
'        ElseIf .RecordCount > 0 Then
'
'            '## 헤더만들기
'            sprStdData.MaxRows = sprStdData.MaxRows + 1
'            sprStdData.Row = sprStdData.MaxRows
'
'            .MoveFirst
'            For ni = 0 To .Fields.Count - 1 Step 1
'                sprStdData.Col = ni + 1
'                sTmp = " ":     If IsNull(.Fields(ni).Name) = False Then sTmp = Trim(.Fields(ni).Name)
'                    Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            Next ni
'
'            .MoveFirst
'            For nRec = 1 To .RecordCount Step 1
'                sprStdData.MaxRows = sprStdData.MaxRows + 1
'                sprStdData.Row = sprStdData.MaxRows
'
'
'                For ni = 0 To .Fields.Count - 1 Step 1
'                    sprStdData.Col = ni + 1
'                    sTmp = " ":     If IsNull(.Fields(ni)) = False Then sTmp = Trim(.Fields(ni))
'                        Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
'                Next ni
'
'                .MoveNext
'
'            Next nRec
'
'
'        End If
'    End With
'
'    nRet = sprStdData.ExportToExcel(sExcelFileName, "Sheet1", sExcelLogFile)
'    MsgBox "엑셀자료 작성완료하였습니다.", vbInformation + vbOKOnly, "전체학생 조회"
'
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    Exit Sub
'
'ErrStmt1:
'    MsgBox "저장할 엑셀명을 등록하세요.", vbExclamation + vbOKOnly, Me.Caption
'    Exit Sub
'
'ErrStmt:
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    MsgBox "전체학생 조회시 에러가 발생하였습니다." & vbCrLf & _
'           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "전체학생 조회"
'
'    On Error GoTo 0
'End Sub
'







Private Sub txtStdNM_F_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtStdNM_F.Text) > " " Then
            Call cmdFind_Click
        End If
    End If
End Sub


Private Sub fpBirth_ymd_F_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(fpBirth_ymd_F.UnFmtText) > " " Then
            Call cmdFind_Click
        End If
    End If

End Sub









'## 결재내역 변경
Private Sub cmdPayChg_Click()
    
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "학생을 조회하세요.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    txtPay.Text = ""
    txtPayChg.Text = ""
    
    With FraPay         '< 결재정보 등록 : 2010.01.13
        .Top = 2700
        .Left = 5700

        .ZOrder 0
        .Visible = False
    End With
    
    txtPay.Text = txtOrdNo.Text
    
    OptPay1.value = True
    OptPay2.value = False
    
    FraPay.Visible = True
    
    txtPay.SetFocus
    
End Sub


Private Sub OptPay1_Click()
    cboCard.Enabled = True
End Sub

Private Sub OptPay2_Click()
    cboCard.Enabled = False
End Sub


Private Sub Label58_Click()
    FraPay.Visible = False
End Sub


Private Sub cmdPaySave_Click()
    '## 등록하기
    Dim sSql        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim nExe        As Integer
    Dim sNo         As String
    
    On Error GoTo Err
    
    
    If Trim(txtPayChg.Text) = "" Then txtPayChg.Text = Trim(txtPay.Text)
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    DBCmd.ActiveConnection = basDataBase.DBConn
    
    basDataBase.DBConn.BeginTrans
    
    '> 수험번호 + 1
    sSql = ""
    sSql = sSql & " UPDATE CLSTD02TB "
    sSql = sSql & "    SET NOW_NUM = NOW_NUM + 1"
    Select Case Trim(basModule.schcd)
        Case "K", "W", "Q"
            sSql = sSql & "  WHERE ACSID   = 'K'"
        Case Else
            sSql = sSql & "  WHERE ACSID   = '" & Trim(basModule.schcd) & "'"
    End Select
    If optExmY.value = True Then
        sSql = sSql & " AND MU_YU = '1'"
    ElseIf optExmN.value = True Then
        sSql = sSql & " AND MU_YU = '0'"
    End If
    sSql = sSql & "     AND KAEYOL= '" & Trim(Right(cboKaeyol.Text, 2)) & "'"

    DBCmd.CommandText = sSql
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe <> 1 Then
        basDataBase.DBConn.RollbackTrans
        GoTo Err
    End If

    '> 결재변경
    sSql = ""
    sSql = sSql & " UPDATE HWPAY01TB "
    sSql = sSql & "    SET ORD_NO = '" & Trim(txtPayChg.Text) & "',"
    sSql = sSql & "        RESULT = '0000',"
    sSql = sSql & "        PAYCONFIRM = 'Y',"
    If OptPay1.value = True Then
        sSql = sSql & "    PAYGUBN = 'C',"
        sSql = sSql & "    DAEPYO  = '" & Trim(Right(cboCard.Text, 4)) & "',"
        sSql = sSql & "    SEPCARD = '" & Trim(Right(cboCard.Text, 4)) & "',"
    ElseIf OptPay2.value = True Then
        sSql = sSql & "    PAYGUBN = 'M',"
    End If
    sSql = sSql & "        PAY_ACCTDATE = SYSDATE,"
    sSql = sSql & "        PAYDATE = SYSDATE"
    sSql = sSql & "  WHERE ORD_NO = '" & Trim(txtPay.Text) & "'"
    
    DBCmd.CommandText = sSql
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe <> 1 Then
        basDataBase.DBConn.RollbackTrans
        GoTo Err
    End If
    
    basDataBase.DBConn.CommitTrans
    
    
    basDataBase.DBConn.BeginTrans
    
        sSql = ""
        sSql = sSql & " SELECT TO_NUMBER(NOW_NUM)-1 AS TN"
        sSql = sSql & "   FROM CLSTD02TB"
        Select Case Trim(basModule.schcd)
            Case "K", "W", "Q"
                sSql = sSql & "  WHERE ACSID   = 'K'"
            Case Else
                sSql = sSql & "  WHERE ACSID   = '" & Trim(basModule.schcd) & "'"
        End Select
        If optExmY.value = True Then
            sSql = sSql & " AND MU_YU = '1'"
        ElseIf optExmN.value = True Then
            sSql = sSql & " AND MU_YU = '0'"
        End If
        sSql = sSql & "     AND KAEYOL= '" & Trim(Right(cboKaeyol.Text, 2)) & "'"
        
        DBCmd.CommandText = sSql
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        sNo = ""
        If DBRec.RecordCount = 1 Then
            DBRec.MoveFirst
            sNo = Trim(DBRec.Fields("TN"))
        End If
            
        
    '> 결재변경
    sSql = ""
    sSql = sSql & " UPDATE CLSTD01TB "
    sSql = sSql & "    SET EXMID  = '" & sNo & "',"
    sSql = sSql & "        ORD_NO = '" & Trim(txtPayChg.Text) & "'"
    sSql = sSql & "  WHERE SCHNO  = '" & Trim(txtSchNo.Text) & "'"
    
    DBCmd.CommandText = sSql
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe <> 1 Then
        basDataBase.DBConn.RollbackTrans
        GoTo Err
    End If
    
    basDataBase.DBConn.CommitTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    MsgBox "등록하였습니다.", vbInformation + vbOKOnly, Me.Caption
    
    Exit Sub
Err:
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    MsgBox "등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, Me.Caption
    
End Sub



































