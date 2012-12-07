VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form STD031 
   Caption         =   "입학사정 >> 등록금 및 가상계좌 부여"
   ClientHeight    =   10395
   ClientLeft      =   3360
   ClientTop       =   3360
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   15240
   Begin VB.Frame Frame4 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame4"
      Height          =   8415
      Left            =   30
      TabIndex        =   20
      Top             =   1170
      Width           =   15045
      Begin VB.Frame Frame5 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame5"
         Height          =   8355
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   14985
         Begin VB.TextBox Text1 
            Height          =   4695
            Left            =   3360
            TabIndex        =   79
            Text            =   "Text1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   7215
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame6"
            Height          =   525
            Left            =   30
            TabIndex        =   63
            Top             =   0
            Width           =   9675
            Begin VB.CommandButton cmdSort 
               Caption         =   "정렬"
               Height          =   375
               Left            =   2010
               TabIndex        =   64
               Top             =   90
               Width           =   645
            End
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   0
               Left            =   2820
               TabIndex        =   65
               Top             =   210
               Width           =   585
               _Version        =   196608
               _ExtentX        =   1032
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
               MaxValue        =   "3"
               MinValue        =   "1"
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
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   1
               Left            =   3480
               TabIndex        =   66
               Top             =   210
               Width           =   585
               _Version        =   196608
               _ExtentX        =   1032
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
               MaxValue        =   "3"
               MinValue        =   "1"
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
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   2
               Left            =   4110
               TabIndex        =   67
               Top             =   210
               Width           =   585
               _Version        =   196608
               _ExtentX        =   1032
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
               MaxValue        =   "3"
               MinValue        =   "1"
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
            Begin VB.Label Label11 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "계열"
               Height          =   210
               Left            =   4080
               TabIndex        =   71
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label8 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "성명"
               Height          =   210
               Left            =   3540
               TabIndex        =   70
               Top             =   15
               Width           =   405
            End
            Begin VB.Label Label7 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "수험번호"
               Height          =   210
               Left            =   2700
               TabIndex        =   69
               Top             =   15
               Width           =   765
            End
            Begin VB.Label Label5 
               BackStyle       =   0  '투명
               Caption         =   "> 정렬"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H008080FF&
               Height          =   285
               Left            =   1800
               TabIndex        =   68
               Top             =   180
               Width           =   645
            End
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "학생 수강료 등록 (&S)"
            Height          =   500
            Left            =   12660
            TabIndex        =   23
            Top             =   7680
            Width           =   2000
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "등록"
            Height          =   225
            Left            =   9450
            TabIndex        =   22
            Top             =   480
            Width           =   675
         End
         Begin FPSpread.vaSpread sprTamgu 
            Height          =   7035
            Left            =   2280
            TabIndex        =   62
            Top             =   570
            Width           =   12675
            _Version        =   393216
            _ExtentX        =   22357
            _ExtentY        =   12409
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
            MaxCols         =   48
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            SpreadDesigner  =   "STD031.frx":0000
         End
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   13950
            TabIndex        =   75
            Top             =   120
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
            MaxValue        =   "2147483647"
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
         Begin VB.Frame fraAMT 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '없음
            Caption         =   "Frame8"
            Height          =   7755
            Left            =   30
            TabIndex        =   24
            Top             =   570
            Width           =   2235
            Begin VB.Frame fraSatam 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  '없음
               Caption         =   "Frame9"
               Height          =   4155
               Left            =   30
               TabIndex        =   80
               Top             =   3600
               Width           =   2175
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   2
                  Left            =   930
                  TabIndex        =   82
                  Top             =   540
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   3
                  Left            =   930
                  TabIndex        =   83
                  Top             =   840
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   4
                  Left            =   930
                  TabIndex        =   84
                  Top             =   1140
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   5
                  Left            =   930
                  TabIndex        =   85
                  Top             =   1500
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   6
                  Left            =   930
                  TabIndex        =   86
                  Top             =   1800
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   7
                  Left            =   930
                  TabIndex        =   87
                  Top             =   2100
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   8
                  Left            =   930
                  TabIndex        =   88
                  Top             =   2400
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   9
                  Left            =   930
                  TabIndex        =   89
                  Top             =   2760
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   10
                  Left            =   930
                  TabIndex        =   90
                  Top             =   3090
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   11
                  Left            =   930
                  TabIndex        =   91
                  Top             =   3750
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin VB.Label Label16 
                  BackStyle       =   0  '투명
                  Caption         =   ">사탐 ------"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   285
                  Left            =   150
                  TabIndex        =   103
                  Top             =   90
                  Width           =   1665
               End
               Begin VB.Label Label29 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "사회문화"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   102
                  Top             =   3165
                  Width           =   945
               End
               Begin VB.Label Label27 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "경제"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   101
                  Top             =   2835
                  Width           =   945
               End
               Begin VB.Label Label26 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "법과정치"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   100
                  Top             =   2475
                  Width           =   945
               End
               Begin VB.Label Label25 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "세계사"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   99
                  Top             =   2175
                  Width           =   945
               End
               Begin VB.Label Label23 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "동아시아사"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   98
                  Top             =   1875
                  Width           =   945
               End
               Begin VB.Label Label22 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "세계지리"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   97
                  Top             =   1560
                  Width           =   945
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "한국지리"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   96
                  Top             =   1200
                  Width           =   945
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "한국사"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   95
                  Top             =   915
                  Width           =   945
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "윤리와사상"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   94
                  Top             =   615
                  Width           =   945
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "생활과윤리"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   93
                  Top             =   315
                  Width           =   945
               End
               Begin VB.Label Label12 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "제2외국어"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   92
                  Top             =   3825
                  Width           =   945
               End
            End
            Begin VB.Frame fraGwatam 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  '없음
               Caption         =   "Frame9"
               Height          =   3705
               Left            =   690
               TabIndex        =   46
               Top             =   3600
               Width           =   3705
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   47
                  Top             =   330
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   2
                  Left            =   930
                  TabIndex        =   48
                  Top             =   660
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   3
                  Left            =   930
                  TabIndex        =   49
                  Top             =   990
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   4
                  Left            =   930
                  TabIndex        =   50
                  Top             =   1320
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   5
                  Left            =   930
                  TabIndex        =   51
                  Top             =   1770
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   6
                  Left            =   930
                  TabIndex        =   52
                  Top             =   2100
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   7
                  Left            =   930
                  TabIndex        =   53
                  Top             =   2430
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   8
                  Left            =   930
                  TabIndex        =   54
                  Top             =   2760
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   9
                  Left            =   930
                  TabIndex        =   55
                  Top             =   3120
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin VB.Label Label33 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "생명과학1"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   106
                  Top             =   1065
                  Width           =   885
               End
               Begin VB.Label Label37 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "생명과학2"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   105
                  Top             =   2505
                  Width           =   885
               End
               Begin VB.Label Label34 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "지구과학1"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   104
                  Top             =   1395
                  Width           =   885
               End
               Begin VB.Label Label45 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수리"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   74
                  Top             =   3165
                  Width           =   555
               End
               Begin VB.Label Label31 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "물리 1"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   61
                  Top             =   390
                  Width           =   555
               End
               Begin VB.Label Label32 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "화학 1"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   60
                  Top             =   735
                  Width           =   555
               End
               Begin VB.Label Label35 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "물리 2"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   59
                  Top             =   1845
                  Width           =   555
               End
               Begin VB.Label Label36 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "화학 2"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   58
                  Top             =   2175
                  Width           =   555
               End
               Begin VB.Label Label38 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "지구과학2"
                  Height          =   210
                  Left            =   -60
                  TabIndex        =   57
                  Top             =   2835
                  Width           =   915
               End
               Begin VB.Label Label39 
                  BackStyle       =   0  '투명
                  Caption         =   ">과탐 ------"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00800000&
                  Height          =   285
                  Left            =   210
                  TabIndex        =   56
                  Top             =   90
                  Width           =   1665
               End
            End
            Begin VB.Frame fraBase 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  '없음
               Caption         =   "Frame7"
               Height          =   3525
               Left            =   30
               TabIndex        =   25
               Top             =   30
               Width           =   2175
               Begin VB.CommandButton cmdAmt 
                  Caption         =   "금액등록(&T)"
                  Height          =   360
                  Left            =   900
                  TabIndex        =   26
                  Top             =   0
                  Width           =   1230
               End
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   27
                  Top             =   360
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   2
                  Left            =   930
                  TabIndex        =   28
                  Top             =   690
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   3
                  Left            =   930
                  TabIndex        =   29
                  Top             =   1020
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   4
                  Left            =   930
                  TabIndex        =   30
                  Top             =   1350
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   270
                  Index           =   8
                  Left            =   930
                  TabIndex        =   36
                  Top             =   3240
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
                  _ExtentY        =   476
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   270
                  Index           =   5
                  Left            =   930
                  TabIndex        =   33
                  Top             =   2430
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
                  _ExtentY        =   476
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   270
                  Index           =   6
                  Left            =   930
                  TabIndex        =   34
                  Top             =   2700
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
                  _ExtentY        =   476
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   270
                  Index           =   7
                  Left            =   930
                  TabIndex        =   35
                  Top             =   2970
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
                  _ExtentY        =   476
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   9
                  Left            =   930
                  TabIndex        =   31
                  Top             =   1710
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   10
                  Left            =   930
                  TabIndex        =   32
                  Top             =   2010
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Text            =   "123,456,789"
                  DecimalPlaces   =   -1
                  DecimalPoint    =   ""
                  FixedPoint      =   0   'False
                  LeadZero        =   0
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin VB.Label Label47 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "기타"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   77
                  Top             =   2085
                  Width           =   945
               End
               Begin VB.Label Label10 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "수익자부담"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   73
                  Top             =   1755
                  Width           =   945
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "등록금"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   45
                  Top             =   435
                  Width           =   555
               End
               Begin VB.Label Label13 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "교재비"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   44
                  Top             =   735
                  Width           =   555
               End
               Begin VB.Label Label14 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "부교재대"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   43
                  Top             =   1065
                  Width           =   765
               End
               Begin VB.Label Label15 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "종합회원비"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   42
                  Top             =   1395
                  Width           =   945
               End
               Begin VB.Label Label40 
                  BackStyle       =   0  '투명
                  Caption         =   ">금액 ------"
                  BeginProperty Font 
                     Name            =   "굴림"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Left            =   30
                  TabIndex        =   41
                  Top             =   120
                  Width           =   1665
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "논술1"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   40
                  Top             =   2475
                  Width           =   945
               End
               Begin VB.Label Label42 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "논술2"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   39
                  Top             =   2775
                  Width           =   945
               End
               Begin VB.Label Label43 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "논술3"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   38
                  Top             =   3045
                  Width           =   945
               End
               Begin VB.Label Label44 
                  Alignment       =   1  '오른쪽 맞춤
                  BackStyle       =   0  '투명
                  Caption         =   "논술4"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   37
                  Top             =   3285
                  Width           =   945
               End
            End
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '투명
            Caption         =   "조회인원"
            Height          =   210
            Left            =   13110
            TabIndex        =   76
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   1035
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14985
         Begin VB.CheckBox chkMusi 
            BackColor       =   &H00D2EAF5&
            Caption         =   "최종접수내역 무시"
            Height          =   255
            Left            =   12750
            TabIndex        =   78
            Top             =   750
            Width           =   2115
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회하기 (&F)"
            Height          =   450
            Left            =   480
            TabIndex        =   6
            Top             =   30
            Width           =   1365
         End
         Begin VB.ComboBox cboHakwon 
            Height          =   300
            Left            =   3660
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   165
            Width           =   975
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   7440
            TabIndex        =   4
            Text            =   "txtStdNM"
            Top             =   578
            Width           =   1005
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3660
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   600
            Width           =   1755
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   5310
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   165
            Width           =   1005
         End
         Begin EditLib.fpMask fpBirth_ymd 
            Height          =   345
            Left            =   9600
            TabIndex        =   7
            Top             =   585
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
         Begin EditLib.fpMask fpExmID_S 
            Height          =   345
            Left            =   7470
            TabIndex        =   8
            Top             =   150
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
         Begin EditLib.fpMask fpExmID_E 
            Height          =   345
            Left            =   8580
            TabIndex        =   9
            Top             =   150
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00D2EAF5&
            Height          =   555
            Left            =   120
            TabIndex        =   10
            Top             =   450
            Width           =   2955
            Begin VB.OptionButton optOkN 
               BackColor       =   &H00D2EAF5&
               Caption         =   "등록금 부여할 학생"
               Height          =   375
               Left            =   30
               TabIndex        =   12
               Top             =   150
               Width           =   1485
            End
            Begin VB.OptionButton optOkY 
               BackColor       =   &H00D2EAF5&
               Caption         =   "부여한 학생"
               Height          =   285
               Left            =   1620
               TabIndex        =   11
               Top             =   180
               Width           =   1275
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "합격학원"
            Height          =   210
            Left            =   2640
            TabIndex        =   19
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "수험번호             부터             까지"
            Height          =   210
            Left            =   6720
            TabIndex        =   18
            Top             =   210
            Width           =   3405
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   8580
            TabIndex        =   17
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            Height          =   210
            Left            =   6420
            TabIndex        =   16
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계  열"
            Height          =   210
            Left            =   2640
            TabIndex        =   15
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회항목"
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
            Left            =   150
            TabIndex        =   14
            Top             =   150
            Width           =   2625
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시험"
            Height          =   210
            Left            =   4890
            TabIndex        =   13
            Top             =   210
            Width           =   375
         End
      End
   End
   Begin VB.Label Label18 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "기본 1"
      Height          =   210
      Left            =   -30
      TabIndex        =   72
      Top             =   75
      Width           =   555
   End
End
Attribute VB_Name = "STD031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD031
'   모 듈  목 적 : 등록금 및 가상계좌 부여 -CP
'
'   작   성   일 : 2007/12/21
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 : 2007.12.21
'   2. 내  용 : 항목추가 : 제2선택 및 기본항목1개
'################################################################################################################

Option Explicit

Private sini_Path       As String        '>> 대성학원
Private sChasuTimes     As String


Private Sub Form_Terminate()
    Unload Me
End Sub


Private Sub Form_Load()
    Dim sData       As String * 255
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sBaseiN     As String           '<< 인문계-기본금액
    Dim sBaseJa     As String           '<< 자연계-기본금액
    
    Dim sBase03     As String           '< 노량진 계열별 처리
    Dim sBase04     As String
    Dim sBase05     As String
    Dim sBase06     As String
    
    Dim sBase07     As String           '< 추가요청 : 2008.05.29
    Dim sBase08     As String
    Dim sBase09     As String
    Dim sBase10     As String
    
    Dim sBase11     As String           '< 추가요청 : 2008.05.30
    Dim sBase12     As String
    Dim sBase13     As String
    Dim sBase14     As String
    Dim sBase15     As String
    Dim sBase16     As String
    
    Dim sSatam      As String           '<< 사탐금액
    
    Dim sGwatam     As String           '<< 과탐금액
    Dim sGwatamNA   As String           '< 수리나형
    
    Dim sSort       As String           '<< sort
    
    Dim sTmp        As String
    
    
    Me.Move 0, 0, 15255, 9980
    
    
    sini_Path = App.Path & "\DAESUNG.INI"       '<< ini file
    sTmp = ""
    nRtn = basModule.GetPrivateProfileString("CHASU", "TIMES", "", sData, 255, sini_Path)
    If nRtn > 0 Then
        sChasuTimes = Left(sData, nRtn)
    Else
        sTmp = "2011011109"
        nRtn = basModule.WritePrivateProfileString("CHASU", "TIMES", sTmp, sini_Path)
        sChasuTimes = sTmp
    End If
    
    
    Me.Tag = "LOAD"
    
        chkMusi.value = 0
        
        fraBase.Move 30, 30, fraAMT.Width - 60, fraBase.Height
        fraSatam.Move 30, 30 + fraBase.Height + 15, fraAMT.Width - 60, fraAMT.Height - fraBase.Height - 75:     fraSatam.Visible = False
        fraGwatam.Move 30, 30 + fraBase.Height + 15, fraAMT.Width - 60, fraAMT.Height - fraBase.Height - 75:    fraGwatam.Visible = False
        
        With sprTamgu
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            
            .Tag = "0"      '<< 다중선택
        End With
        
        With cboHakwon
            .Clear
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            
            Select Case basModule.SchCD
                Case "N"
                    .ListIndex = 0
                Case "K"
                    .ListIndex = 1
                Case "S"
                    .ListIndex = 2
                Case "P"
                    .ListIndex = 3
                Case "M"
                    .ListIndex = 4
                    
                Case "W"
                    .ListIndex = 5
                Case "Q"
                    .ListIndex = 6
                    
                Case "J"
                    .ListIndex = 7
                Case "B"
                    .ListIndex = 8
                    
            End Select
        End With
        
        With cboExmType
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "무시험" & Space(30) & "0"
            .AddItem "유시험" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        
'>> 계열
        With cboKaeyol
            .Clear
            
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            
        '<< 계열 >> : 2008.01.09
            If Trim(basModule.SchCD) = "N" Then             '< 노량진
                .AddItem "예체" & Space(30) & "03"
                .AddItem "수리(나)" & Space(30) & "04"
                .AddItem "인문수능" & Space(30) & "05"
                .AddItem "자연수능" & Space(30) & "06"
                
                .AddItem "인문-신" & Space(30) & "07"
                .AddItem "자연-신" & Space(30) & "08"
                '.AddItem "수능인문-신" & Space(30) & "09"
                '.AddItem "수능자연-신" & Space(30) & "10"
                
                .AddItem "편)인문" & Space(30) & "11"
                .AddItem "편)자연" & Space(30) & "12"
                .AddItem "편)예체" & Space(30) & "13"
                .AddItem "편)수리(나)" & Space(30) & "14"
                .AddItem "편)인문수능" & Space(30) & "15"
                .AddItem "편)자연수능" & Space(30) & "16"
            End If
        '<< 계열 >> : 2008.01.10
            If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Or Trim(basModule.SchCD) = "M" Then        '< 강남 2008.03.24
                .AddItem "주말법대" & Space(30) & "04"
                .AddItem "주말의대" & Space(30) & "05"
                
                .AddItem "야간법대" & Space(30) & "06"
                .AddItem "야간의대" & Space(30) & "07"
            
                .AddItem "선착순인문" & Space(30) & "11"
                .AddItem "선착순자연" & Space(30) & "12"
                
                .AddItem "선착순인문16" & Space(30) & "16"
                .AddItem "선착순자연17" & Space(30) & "17"
                
                .AddItem "내신우수자인문" & Space(30) & "19"
                .AddItem "내신우수자자연" & Space(30) & "20"
                
            End If
        '<< 계열 >> : 2008.02.15
            If Trim(basModule.SchCD) = "S" Then             '< 송파
''                .AddItem "예체능" & Space(30) & "03"
''
''                .AddItem "인문수능" & Space(30) & "05"
''                .AddItem "자연수능" & Space(30) & "06"
''
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
                
                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"
                
                .AddItem "서울대특별인문" & Space(30) & "21"
                .AddItem "서울대특별자연" & Space(30) & "22"
                .AddItem "야간서울대인문" & Space(30) & "23"
                .AddItem "야간서울대자연" & Space(30) & "24"
             
                
            End If
            
        '<< 계열 >> : 2008.02.15
            If Trim(basModule.SchCD) = "P" Then             '< 마송
                .AddItem "특별인문" & Space(30) & "03"
                .AddItem "특별자연" & Space(30) & "04"
                
            End If
            
            If Trim(basModule.SchCD) = "J" Then             '< 양재
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
                
                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"
                
            End If
            
        '<< 계열 >> : 2009.01.09
            If Trim(basModule.SchCD) = "B" Then             '< 부산
                .AddItem "수학선행인문" & Space(30) & "05"
                .AddItem "수학선행자연" & Space(30) & "06"
                
                .AddItem "연.고대인문" & Space(30) & "07"
                .AddItem "연.고대자연" & Space(30) & "08"
                
                .AddItem "심화인문" & Space(30) & "09"
                .AddItem "심화자연" & Space(30) & "10"
                
            End If
            
            
            .AddItem "계열무시" & Space(30) & "98"
            '.AddItem "계열무시자연" & Space(30) & "99"
            
            
            .ListIndex = 0
        End With
        
        
        sini_Path = App.Path & "\DAESUNG.INI"
        
        '>> 프로그램 INI 파일
        If Dir(sini_Path) = "" Then                                     '<< 파일이 없으면 생성
            sBaseiN = insert_AMT_ini_File("BASEIN", "0/0/0/0/0/0/0/0/0/0/")         '< 인문계
            sBaseJa = insert_AMT_ini_File("BASEJA", "0/0/0/0/0/0/0/0/0/0/")         '< 자연계
            
            sBase03 = insert_AMT_ini_File("BASE03", "0/0/0/0/0/0/0/0/0/0/")         '< 노량진 : 계열별
            sBase04 = insert_AMT_ini_File("BASE04", "0/0/0/0/0/0/0/0/0/0/")
            sBase05 = insert_AMT_ini_File("BASE05", "0/0/0/0/0/0/0/0/0/0/")
            sBase06 = insert_AMT_ini_File("BASE06", "0/0/0/0/0/0/0/0/0/0/")
            
            sBase07 = insert_AMT_ini_File("BASE07", "0/0/0/0/0/0/0/0/0/0/")
            sBase08 = insert_AMT_ini_File("BASE08", "0/0/0/0/0/0/0/0/0/0/")
            sBase09 = insert_AMT_ini_File("BASE09", "0/0/0/0/0/0/0/0/0/0/")
            sBase10 = insert_AMT_ini_File("BASE10", "0/0/0/0/0/0/0/0/0/0/")
            
            sBase11 = insert_AMT_ini_File("BASE11", "0/0/0/0/0/0/0/0/0/0/")
            sBase12 = insert_AMT_ini_File("BASE12", "0/0/0/0/0/0/0/0/0/0/")
            sBase13 = insert_AMT_ini_File("BASE13", "0/0/0/0/0/0/0/0/0/0/")
            sBase14 = insert_AMT_ini_File("BASE14", "0/0/0/0/0/0/0/0/0/0/")
            sBase15 = insert_AMT_ini_File("BASE15", "0/0/0/0/0/0/0/0/0/0/")
            sBase16 = insert_AMT_ini_File("BASE16", "0/0/0/0/0/0/0/0/0/0/")
            
            sSatam = insert_AMT_ini_File("SATAM", "0/0/0/0/0/0/0/0/0/0/0/0/")
            sGwatam = insert_AMT_ini_File("GWATAM", "0/0/0/0/0/0/0/0/0/")
        End If
        
        sGbn = "STD031"
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASEIN", "", sData, 255, sini_Path)         '>> 인문계-기본금액
            sBaseiN = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBaseiN = insert_AMT_ini_File("BASEIN", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASEJA", "", sData, 255, sini_Path)         '>> 인문계-기본금액
            sBaseJa = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBaseJa = insert_AMT_ini_File("BASEJA", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            '>> 노량진 계열별
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE03", "", sData, 255, sini_Path)         '>> 03
            sBase03 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase03 = insert_AMT_ini_File("BASE03", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE04", "", sData, 255, sini_Path)         '>> 04
            sBase04 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase04 = insert_AMT_ini_File("BASE04", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE05", "", sData, 255, sini_Path)         '>> 05
            sBase05 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase05 = insert_AMT_ini_File("BASE05", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE06", "", sData, 255, sini_Path)         '>> 06
            sBase06 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase06 = insert_AMT_ini_File("BASE06", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE07", "", sData, 255, sini_Path)         '>> 07
            sBase07 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase07 = insert_AMT_ini_File("BASE07", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE08", "", sData, 255, sini_Path)         '>> 08
            sBase08 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase08 = insert_AMT_ini_File("BASE08", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE09", "", sData, 255, sini_Path)         '>> 09
            sBase09 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase09 = insert_AMT_ini_File("BASE09", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE10", "", sData, 255, sini_Path)         '>> 10
            sBase10 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase10 = insert_AMT_ini_File("BASE10", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE11", "", sData, 255, sini_Path)         '>> 11
            sBase11 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase11 = insert_AMT_ini_File("BASE11", "0/0/0/0/0/0/0/0/0/0/")
            End If
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE12", "", sData, 255, sini_Path)         '>> 12
            sBase12 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase12 = insert_AMT_ini_File("BASE12", "0/0/0/0/0/0/0/0/0/0/")
            End If
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE13", "", sData, 255, sini_Path)         '>> 13
            sBase13 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase13 = insert_AMT_ini_File("BASE13", "0/0/0/0/0/0/0/0/0/0/")
            End If
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE14", "", sData, 255, sini_Path)         '>> 14
            sBase14 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase14 = insert_AMT_ini_File("BASE14", "0/0/0/0/0/0/0/0/0/0/")
            End If
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE15", "", sData, 255, sini_Path)         '>> 15
            sBase15 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase15 = insert_AMT_ini_File("BASE15", "0/0/0/0/0/0/0/0/0/0/")
            End If
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE16", "", sData, 255, sini_Path)         '>> 16
            sBase16 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase16 = insert_AMT_ini_File("BASE16", "0/0/0/0/0/0/0/0/0/0/")
            End If
            
            '## ---------------------------------- ##
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SATAM", "", sData, 255, sini_Path)          '>> 사탐금액
            sSatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSatam = insert_AMT_ini_File("SATAM", "0/0/0/0/0/0/0/0/0/0/0/0/")
            End If
            
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "GWATAM", "", sData, 255, sini_Path)         '>> 과탐금액
            sGwatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sGwatam = insert_AMT_ini_File("GWATAM", "0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "GWATAMNA", "", sData, 255, sini_Path)       '>> 사탐금액 : 수리나형
            sGwatamNA = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSatam = insert_AMT_ini_File("GWATAMNA", "0/0/0/0/0/0/0/0/0/")
            End If
            
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SORT", "", sData, 255, sini_Path)           '>> SORT 순서
            sSort = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSort = insert_AMT_ini_File("SORT", "0,3/1,2/2,1/")
            End If
            
        Call init_Form(sBaseiN, sBaseJa, sBase03, sBase04, sBase05, sBase06, sBase07, sBase08, sBase09, sBase10, sBase11, sBase12, sBase13, sBase14, sBase15, sBase16, sSatam, sGwatam, sGwatamNA, sSort)
        
    Me.Tag = ""
End Sub



'/* 신규처리 */
Private Sub init_Form(ByVal aBaseiN As String, _
                      ByVal aBaseJa As String, _
                      ByVal aBase03 As String, _
                      ByVal aBase04 As String, _
                      ByVal aBase05 As String, _
                      ByVal aBase06 As String, _
                      ByVal aBase07 As String, _
                      ByVal aBase08 As String, _
                      ByVal aBase09 As String, _
                      ByVal aBase10 As String, _
                      ByVal aBase11 As String, _
                      ByVal aBase12 As String, _
                      ByVal aBase13 As String, _
                      ByVal aBase14 As String, _
                      ByVal aBase15 As String, _
                      ByVal aBase16 As String, _
                      ByVal aSatam As String, _
                      ByVal aGwatam As String, _
                      ByVal aGwatamNA As String, _
                      ByVal aSort As String)
                      
    Dim ni      As Integer
    Dim sDivs() As String
    Dim sDivC() As String
    
    fpTotCnt.value = 0
    
    optOkN.value = True         '< 등록금 부여할 학생
    optOkY.value = False        '< 부여한 학생

    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""

    sprTamgu.MaxRows = 0        '< spread
    
    
'>> 금액 정리
    For ni = 1 To 10 Step 1     '< base 금액 index
        fpBase(ni).value = 0
    Next ni

    '사탐 satam
    For ni = 1 To SATAM_COUNT + 1 Step 1
        fpSatam(ni).value = 0
    Next ni
    
    For ni = 1 To 9 Step 1
        fpGwatam(ni).value = 0
    Next ni

    '>> 금액 공통
    Select Case Trim(basModule.SchCD)
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "02"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                    
                '>> 계열무시
                Case "98"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "99"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                    
                    
                '> 노량진 각 계열별로
                Case "03"
                    sDivs() = Split(aBase03, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "04"
                    sDivs() = Split(aBase04, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "05"
                    sDivs() = Split(aBase05, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "06"
                    sDivs() = Split(aBase06, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "07"
                    sDivs() = Split(aBase07, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "08"
                    sDivs() = Split(aBase08, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "09"
                    sDivs() = Split(aBase09, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "10"
                    sDivs() = Split(aBase10, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                
                Case "11"
                    sDivs() = Split(aBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(aBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "13"
                    sDivs() = Split(aBase13, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "14"
                    sDivs() = Split(aBase14, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "15"
                    sDivs() = Split(aBase15, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "16"
                    sDivs() = Split(aBase16, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                '## ---------------------- ##
                    
            End Select
        Case "K", "W", "Q", "M"             '< 계열 : 2008.01.10 : 강남
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "04", "06", "11", "16"             '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "05", "07", "12", "17"             '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                '>> 계열무시
                Case "98"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "99"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
            End Select
            
            
        Case "S", "P"           '< 계열 : 2008.02.15 : 송파/ 마송
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"                         '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "04"                         '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                Case "05"
                    sDivs() = Split(aBase05, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "06"
                    sDivs() = Split(aBase06, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni


                Case "11"
                    sDivs() = Split(aBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(aBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                Case "18"
                    sDivs() = Split(aBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "19"
                    sDivs() = Split(aBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                '>> 계열무시
                Case "98"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "99"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni

                
            End Select
            
        Case "J"                                        '< 계열 : 양재
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"                         '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "04"                         '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                Case "11"
                    sDivs() = Split(aBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(aBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "18"
                    sDivs() = Split(aBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "19"
                    sDivs() = Split(aBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                
                
                '>> 계열무시
                Case "98"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "99"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni

                
            End Select
            
            
        Case Else
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"             '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02"                   '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                '>> 계열무시
                Case "98"           '< 인문계
                    sDivs() = Split(aBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "99"           '< 자연계
                    sDivs() = Split(aBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
            End Select
    End Select
    
    
    '>> 사탐
    sDivs() = Split(aSatam, "/", -1, vbTextCompare)
    For ni = 0 To UBound(sDivs) - 1 Step 1
        fpSatam(ni + 1).value = CLng(sDivs(ni))
    Next ni
    
    '>> 과탐
    Select Case Trim(basModule.SchCD)
        Case "N"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "04"           '< 수리나형
                    sDivs() = Split(aGwatamNA, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpGwatam(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case Else
                    sDivs() = Split(aGwatam, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpGwatam(ni + 1).value = CLng(sDivs(ni))
                    Next ni
            End Select
        Case Else
            sDivs() = Split(aGwatam, "/", -1, vbTextCompare)
            For ni = 0 To UBound(sDivs) - 1 Step 1
                fpGwatam(ni + 1).value = CLng(sDivs(ni))
            Next ni
    End Select
    
    
    '>> sort
        sDivs() = Split(aSort, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
            
            fpSort(CInt(sDivC(0))).value = CInt(sDivC(1))
        Next ni
    
End Sub




'>> 콤보 계열내역 선택시
Private Sub cboKaeyol_Click()
    
    Dim ni      As Integer
    Dim sDivs() As String
    
    Dim sData       As String * 255
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sBaseiN     As String           '<< 인문계-기본금액
    Dim sBaseJa     As String           '<< 자연계-기본금액
    
    Dim sBase03     As String           '< 노량진 계열별 처리
    Dim sBase04     As String
    Dim sBase05     As String
    Dim sBase06     As String
    
    Dim sBase07     As String
    Dim sBase08     As String
    Dim sBase09     As String
    Dim sBase10     As String
    
    Dim sBase11     As String
    Dim sBase12     As String
    Dim sBase13     As String
    Dim sBase14     As String
    Dim sBase15     As String
    Dim sBase16     As String
    
    Dim sSatam      As String           '<< 사탐금액
    
    Dim sGwatam     As String           '<< 과탐금액
    Dim sGwatamNA    As String          '<  과탐 나형
    
    
'ini 파일 불러오기
    sGbn = "STD031"
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASEIN", "", sData, 255, sini_Path)         '>> 인문계-기본금액
        sBaseiN = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBaseiN = insert_AMT_ini_File("BASEIN", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASEJA", "", sData, 255, sini_Path)         '>> 인문계-기본금액
        sBaseJa = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBaseJa = insert_AMT_ini_File("BASEJA", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
         '>> 노량진 계열별
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE03", "", sData, 255, sini_Path)         '>> 03
        sBase03 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase03 = insert_AMT_ini_File("BASE03", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE04", "", sData, 255, sini_Path)         '>> 04
        sBase04 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase04 = insert_AMT_ini_File("BASE04", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE05", "", sData, 255, sini_Path)         '>> 05
        sBase05 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase05 = insert_AMT_ini_File("BASE05", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE06", "", sData, 255, sini_Path)         '>> 06
        sBase06 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase06 = insert_AMT_ini_File("BASE06", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE07", "", sData, 255, sini_Path)         '>> 07
        sBase07 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase07 = insert_AMT_ini_File("BASE07", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE08", "", sData, 255, sini_Path)         '>> 08
        sBase08 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase08 = insert_AMT_ini_File("BASE08", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE09", "", sData, 255, sini_Path)         '>> 09
        sBase09 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase09 = insert_AMT_ini_File("BASE09", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE10", "", sData, 255, sini_Path)         '>> 10
        sBase10 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase10 = insert_AMT_ini_File("BASE10", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE11", "", sData, 255, sini_Path)         '>> 11
        sBase11 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase11 = insert_AMT_ini_File("BASE11", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE12", "", sData, 255, sini_Path)         '>> 12
        sBase12 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase12 = insert_AMT_ini_File("BASE12", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE13", "", sData, 255, sini_Path)         '>> 13
        sBase13 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase13 = insert_AMT_ini_File("BASE13", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE14", "", sData, 255, sini_Path)         '>> 14
        sBase14 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase14 = insert_AMT_ini_File("BASE14", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE15", "", sData, 255, sini_Path)         '>> 15
        sBase15 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase15 = insert_AMT_ini_File("BASE15", "0/0/0/0/0/0/0/0/0/0/")
        End If
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "BASE16", "", sData, 255, sini_Path)         '>> 16
        sBase16 = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sBase16 = insert_AMT_ini_File("BASE16", "0/0/0/0/0/0/0/0/0/0/")
        End If
        
        '## ---------------------------------- ##
            
            
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "SATAM", "", sData, 255, sini_Path)          '>> 사탐금액
        sSatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sSatam = insert_AMT_ini_File("SATAM", "0/0/0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "GWATAM", "", sData, 255, sini_Path)         '>> 과탐금액
        sGwatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sGwatam = insert_AMT_ini_File("GWATAM", "0/0/0/0/0/0/0/0/0/")
        End If
        
        sData = ""
        nRtn = basModule.GetPrivateProfileString(sGbn, "GWATAMNA", "", sData, 255, sini_Path)       '>> 수리나형
        sGwatamNA = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
        If nRtn = 0 Then
            sGwatamNA = insert_AMT_ini_File("GWATAMNA", "0/0/0/0/0/0/0/0/0/")
        End If
        
    
'금액내역 보여주기
    For ni = 1 To 10 Step 1
        fpBase(ni).value = 0
    Next ni
    
    '사탐
    For ni = 1 To SATAM_COUNT + 1 Step 1
        fpSatam(ni).value = 0
    Next ni
    
    
    For ni = 1 To 9 Step 1
        fpGwatam(ni).value = 0
    Next ni

    '>> 금액
    Select Case Trim(basModule.SchCD)
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                    
                Case "98"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "99"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
                '>> 노량진 계열별 추가
                Case "03"
                    sDivs() = Split(sBase03, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "04"
                    sDivs() = Split(sBase04, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "05"
                    sDivs() = Split(sBase05, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "06"
                    sDivs() = Split(sBase06, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "07"
                    sDivs() = Split(sBase07, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "08"
                    sDivs() = Split(sBase08, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "09"
                    sDivs() = Split(sBase09, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "10"
                    sDivs() = Split(sBase10, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "11"
                    sDivs() = Split(sBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(sBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "13"
                    sDivs() = Split(sBase13, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "14"
                    sDivs() = Split(sBase14, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "15"
                    sDivs() = Split(sBase15, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "16"
                    sDivs() = Split(sBase16, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                    
            End Select
        Case "K", "W", "Q", "M"                     '< 계열 : 2008.01.10 : 강남
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "04", "06", "11", "16"                        '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "05", "06", "12", "17"                         '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                
                Case "98"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "99"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
            End Select
        Case "S", "P"               '< 계열 : 2008.02.15 : 송파
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"                                     '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "04"                                     '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "05"
                    sDivs() = Split(sBase05, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "06"
                    sDivs() = Split(sBase06, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                
                
                Case "11"
                    sDivs() = Split(sBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(sBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "18"
                    sDivs() = Split(sBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "19"
                    sDivs() = Split(sBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni

                
                Case "98"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "99"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
            End Select
            
            
        Case "J"               '< 계열 : 양재
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"                                     '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02", "04"                                     '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "11"
                    sDivs() = Split(sBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "12"
                    sDivs() = Split(sBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "18"
                    sDivs() = Split(sBase11, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "19"
                    sDivs() = Split(sBase12, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                
                Case "98"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "99"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    

            End Select
            
        Case Else
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"                                 '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "02"                                       '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
                Case "98"                   '< 인문계
                    sDivs() = Split(sBaseiN, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case "99"                   '< 자연계
                    sDivs() = Split(sBaseJa, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpBase(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                    
            End Select
    End Select
    
    '>> 사탐
        sDivs() = Split(sSatam, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            fpSatam(ni + 1).value = CLng(sDivs(ni))
        Next ni

    
    '>> 과탐
    Select Case Trim(basModule.SchCD)
        Case "N"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "04"                       '< 수리나형
                    sDivs() = Split(sGwatamNA, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpGwatam(ni + 1).value = CLng(sDivs(ni))
                    Next ni
                Case Else
                    sDivs() = Split(sGwatam, "/", -1, vbTextCompare)
                    For ni = 0 To UBound(sDivs) - 1 Step 1
                        fpGwatam(ni + 1).value = CLng(sDivs(ni))
                    Next ni
            End Select
        Case Else
            sDivs() = Split(sGwatam, "/", -1, vbTextCompare)
            For ni = 0 To UBound(sDivs) - 1 Step 1
                fpGwatam(ni + 1).value = CLng(sDivs(ni))
            Next ni
    End Select
    
    
    Select Case Trim(basModule.SchCD)           '< 2008.01.09
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "05", "07", "09", "11", "13", "15", "98"
                    fraSatam.Visible = True
                    fraGwatam.Visible = False
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "인문 선택영역 내용"
                            .Col = 23:          .Text = "인문 논술선택"
                            .Col = 37:          .Text = "인문 선택영역 금액내용"        '< 변경
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:          .Text = constSatams(0)
                            .Col = 12:          .Text = constSatams(1)
                            .Col = 13:          .Text = constSatams(2)
                            .Col = 14:          .Text = constSatams(3)
                            .Col = 15:          .Text = constSatams(4)
                            .Col = 16:          .Text = constSatams(5)
                            .Col = 17:          .Text = constSatams(6)
                            .Col = 18:          .Text = constSatams(7)
                            .Col = 19:          .Text = constSatams(8)
                            .Col = 20:          .Text = constSatams(9)
                            .Col = 21:          .Text = ""
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"            '변경
                            .Col = 26:          .Text = ""                  '변경
                            
                            .Col = 37:          .Text = constSatams(0)
                            .Col = 38:          .Text = constSatams(1)
                            .Col = 39:          .Text = constSatams(2)
                            .Col = 40:          .Text = constSatams(3)
                            .Col = 41:          .Text = constSatams(4)
                            .Col = 42:          .Text = constSatams(5)
                            .Col = 43:          .Text = constSatams(6)
                            .Col = 44:          .Text = constSatams(7)
                            .Col = 45:          .Text = constSatams(8)
                            .Col = 46:          .Text = constSatams(9)
                            .Col = 47:          .Text = ""
                    End With
                    
                Case "02", "04", "06", "08", "10", "12", "14", "16", "99"
                    fraSatam.Visible = False
                    fraGwatam.Visible = True
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "자연 선택영역 내용"
                            .Col = 23:          .Text = "자연 논술선택"
                            .Col = 36:          .Text = "자연 선택영역 금액내용"
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:           .Text = "물1"
                            .Col = 12:          .Text = "화1"
                            .Col = 13:          .Text = "생1"
                            .Col = 14:          .Text = "지1"
                            .Col = 15:          .Text = "물2"
                            .Col = 16:          .Text = "화2"
                            .Col = 17:          .Text = "생2"
                            .Col = 18:          .Text = "지2"
                            .Col = 19:          .Text = "-"
                            .Col = 20:          .Text = "-"
                            .Col = 21:          .Text = "-"
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"        '변경
                            .Col = 26:          .Text = ""              '변경
                            
                            .Col = 37:          .Text = "물1"
                            .Col = 38:          .Text = "화1"
                            .Col = 39:          .Text = "생1"
                            .Col = 40:          .Text = "지1"
                            .Col = 41:          .Text = "물2"
                            .Col = 42:          .Text = "화2"
                            .Col = 43:          .Text = "생2"
                            .Col = 44:          .Text = "지2"
                            .Col = 45:          .Text = "-"
                            .Col = 46:          .Text = "-"
                            .Col = 47:          .Text = "-"
                            
                    End With
                    
            End Select
        
        Case "S", "P", "J"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "05", "11", "18", "98"                                 '< 계열 : 송파/ 마송 2008.02.15
                    fraSatam.Visible = True
                    fraGwatam.Visible = False
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "인문 선택영역 내용"
                            .Col = 23:          .Text = "인문 논술선택"
                            .Col = 37:          .Text = "인문 선택영역 금액내용"        '< 변경
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:          .Text = constSatams(0)
                            .Col = 12:          .Text = constSatams(1)
                            .Col = 13:          .Text = constSatams(2)
                            .Col = 14:          .Text = constSatams(3)
                            .Col = 15:          .Text = constSatams(4)
                            .Col = 16:          .Text = constSatams(5)
                            .Col = 17:          .Text = constSatams(6)
                            .Col = 18:          .Text = constSatams(7)
                            .Col = 19:          .Text = constSatams(8)
                            .Col = 20:          .Text = constSatams(9)
                            .Col = 21:          .Text = ""
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"            '< 변경
                            .Col = 26:          .Text = ""                  '< 변경
                            
                            .Col = 37:          .Text = constSatams(0)
                            .Col = 38:          .Text = constSatams(1)
                            .Col = 39:          .Text = constSatams(2)
                            .Col = 40:          .Text = constSatams(3)
                            .Col = 41:          .Text = constSatams(4)
                            .Col = 42:          .Text = constSatams(5)
                            .Col = 43:          .Text = constSatams(6)
                            .Col = 44:          .Text = constSatams(7)
                            .Col = 45:          .Text = constSatams(8)
                            .Col = 46:          .Text = constSatams(9)
                            .Col = 47:          .Text = ""
                    End With
                    
                Case "02", "04", "06", "08", "12", "19", "99"                               '< 계열 : 송파 : 2008.02.15
                    fraSatam.Visible = False
                    fraGwatam.Visible = True
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "자연 선택영역 내용"
                            .Col = 23:          .Text = "자연 논술선택"
                            .Col = 36:          .Text = "자연 선택영역 금액내용"
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:           .Text = "물1"
                            .Col = 12:          .Text = "화1"
                            .Col = 13:          .Text = "생1"
                            .Col = 14:          .Text = "지1"
                            .Col = 15:          .Text = "물2"
                            .Col = 16:          .Text = "화2"
                            .Col = 17:          .Text = "생2"
                            .Col = 18:          .Text = "지2"
                            .Col = 19:          .Text = "-"
                            .Col = 20:          .Text = "-"
                            .Col = 21:          .Text = "-"
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"            '< 변경
                            .Col = 26:          .Text = ""                  '< 변경
                            
                            .Col = 37:          .Text = "물1"
                            .Col = 38:          .Text = "화1"
                            .Col = 39:          .Text = "생1"
                            .Col = 40:          .Text = "지1"
                            .Col = 41:          .Text = "물2"
                            .Col = 42:          .Text = "화2"
                            .Col = 43:          .Text = "생2"
                            .Col = 44:          .Text = "지2"
                            .Col = 45:          .Text = "-"
                            .Col = 46:          .Text = "-"
                            .Col = 47:          .Text = "-"
                            
                    End With
                    
            End Select
        
        
        Case Else
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "04", "06", "11", "16", "98"                    '< 계열 : 강남 2008.01.10
                    fraSatam.Visible = True
                    fraGwatam.Visible = False
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "인문 선택영역 내용"
                            .Col = 23:          .Text = "인문 논술선택"
                            .Col = 37:          .Text = "인문 선택영역 금액내용"        '< 변경
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:          .Text = constSatams(0)
                            .Col = 12:          .Text = constSatams(1)
                            .Col = 13:          .Text = constSatams(2)
                            .Col = 14:          .Text = constSatams(3)
                            .Col = 15:          .Text = constSatams(4)
                            .Col = 16:          .Text = constSatams(5)
                            .Col = 17:          .Text = constSatams(6)
                            .Col = 18:          .Text = constSatams(7)
                            .Col = 19:          .Text = constSatams(8)
                            .Col = 20:          .Text = constSatams(9)
                            .Col = 21:          .Text = ""
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"            '< 변경
                            .Col = 26:          .Text = ""                  '< 변경
                            
                            .Col = 37:          .Text = constSatams(0)
                            .Col = 38:          .Text = constSatams(1)
                            .Col = 39:          .Text = constSatams(2)
                            .Col = 40:          .Text = constSatams(3)
                            .Col = 41:          .Text = constSatams(4)
                            .Col = 42:          .Text = constSatams(5)
                            .Col = 43:          .Text = constSatams(6)
                            .Col = 44:          .Text = constSatams(7)
                            .Col = 45:          .Text = constSatams(8)
                            .Col = 46:          .Text = constSatams(9)
                            .Col = 47:          .Text = ""
                    End With
                    
                Case "02", "05", "07", "12", "17", "99"                        '< 계열 : 강남 : 2008.01.10"
                    fraSatam.Visible = False
                    fraGwatam.Visible = True
                    
                    '>> spread header 변경
                    With sprTamgu
                        .Row = SpreadHeader
                            .Col = 11:           .Text = "자연 선택영역 내용"
                            .Col = 23:          .Text = "자연 논술선택"
                            .Col = 36:          .Text = "자연 선택영역 금액내용"
                            
                        .Row = SpreadHeader + 1
                            .Col = 11:           .Text = "물1"
                            .Col = 12:          .Text = "화1"
                            .Col = 13:          .Text = "생1"
                            .Col = 14:          .Text = "지1"
                            .Col = 15:          .Text = "물2"
                            .Col = 16:          .Text = "화2"
                            .Col = 17:          .Text = "생2"
                            .Col = 18:          .Text = "지2"
                            .Col = 19:          .Text = "-"
                            .Col = 20:          .Text = "-"
                            .Col = 21:          .Text = "-"
                            
                            .Col = 22:          .Text = "제2외"
                            
                            .Col = 23:          .Text = "언어"
                            .Col = 24:          .Text = "수리"
                            .Col = 25:          .Text = "외국어"            '< 변경
                            .Col = 26:          .Text = ""                  '< 변경
                            
                            .Col = 37:          .Text = "물1"
                            .Col = 38:          .Text = "화1"
                            .Col = 39:          .Text = "생1"
                            .Col = 40:          .Text = "지1"
                            .Col = 41:          .Text = "물2"
                            .Col = 42:          .Text = "화2"
                            .Col = 43:          .Text = "생2"
                            .Col = 44:          .Text = "지2"
                            .Col = 45:          .Text = "-"
                            .Col = 46:          .Text = "-"
                            .Col = 47:          .Text = "-"
                            
                    End With
                    
            End Select

    End Select
    
End Sub


Private Function insert_AMT_ini_File(ByVal aGbn As String, ByVal aData As String) As String
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sReturn     As String
    
    sGbn = "STD031"
         nRtn = basModule.WritePrivateProfileString(sGbn, aGbn, aData, sini_Path)        '<< 구분자에 따른 금액등록
         
    insert_AMT_ini_File = aData
    
End Function




'>> 합격생 조회
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
    
    chkAll.value = 0
    sprTamgu.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, EXMID, "
    sStr = sStr & "         STDNM,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "             '01'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "             '02'"
    sStr = sStr & "         END END GAEYUL_CD,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "             '사탐'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "             '과탐'"
    sStr = sStr & "         END END GAEYUL,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT, CY_ACNT2 , CY_ACNT3,"
    sStr = sStr & "         TOT_AMT,"
    sStr = sStr & "         0 AS CHKS,"
    sStr = sStr & "  "
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
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* 사탐-윤리사상 */"
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
    sStr = sStr & " '' AS SEL11,"
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
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '42'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '81'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '82'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '83'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '84'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END END END END END END END END END END END END END END END AS SEL_X2,"
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
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"          '< 변경
    sStr = sStr & "             '93'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"                '< 변경
    sStr = sStr & "             '94'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N4,"
    
    '---------------------------------------------------------------------------------------------------------
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS B1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS B2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS B3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS B4  ,"
    
    sStr = sStr & "         NVL(BASE_AMT9    ,0) AS B5  ,"       '< 추가 : 2007.12.21
    sStr = sStr & "         NVL(BASE_AMT10   ,0) AS B6  ,"       '< 추가 : 2008.01.09
    
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS B7  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS B8  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS B9  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS B10 ,"
    '---------------------------------------------------------------------------------------------------------
    
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS TAMGU_AMT1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS TAMGU_AMT2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS TAMGU_AMT3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS TAMGU_AMT4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS TAMGU_AMT5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS TAMGU_AMT6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS TAMGU_AMT7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS TAMGU_AMT8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS TAMGU_AMT9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS TAMGU_AMT10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS TAMGU_AMT11,"
    sStr = sStr & "         NVL(TAMGU_AMT12  ,0) AS TAMGU_AMT12"        '< 추가 : 2007.12.21
    
    sStr = sStr & "  "
    sStr = sStr & "    FROM CLSTD01TB"
    'sStr = sStr & "    WHERE TOT_AMT = 0"
    sStr = sStr & "   WHERE (PASS1 = ? OR"
    sStr = sStr & "          PASS2 = ? OR"
    sStr = sStr & "          PASS3 = ? OR"
    sStr = sStr & "          PASS4 = ? )"
'>> 등록금 등록여부
    If optOkN.value = True Then
        sStr = sStr & " AND TOT_AMT = 0 "
    ElseIf optOkY.value = True Then
        sStr = sStr & " AND TOT_AMT > 0 "
    End If
'>> 시험
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "ALL"
            ' NO ACTION
        Case "0"
            sStr = sStr & " AND EXMTYPE = '0' "     '<< 무시험
        Case "1"
            sStr = sStr & " AND EXMTYPE = '1' "     '<< 유시험
    End Select
'>> 계열
'    Select Case Trim(Right(cboKaeyol, 30))
'        Case "XX"
'            ' no action
'        Case "01", "03"
'            sStr = sStr & " AND SEL1 > ' ' "
'        Case "02"
'            sStr = sStr & " AND SEL3 > ' ' "
'    End Select
    
    If Trim(Right(cboKaeyol.Text, 30)) = "98" Then
        'NO Action
    ElseIf Trim(Right(cboKaeyol.Text, 30)) = "99" Then
        'NO Action
    Else
        sStr = sStr & " AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    End If
    

'>> 수험번호
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
'>> 학생명
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & " AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
    End If
'>> 생년월일
    If Trim(fpBirth_ymd.UnFmtText) > " " Then
        sStr = sStr & " AND Birth_ymd LIKE '" & Trim(fpBirth_ymd.UnFmtText) & "%'"
    End If
    
    
    
    Select Case basModule.SchCD
        Case "K"
            sStr = sStr & "         AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            
        Case Else
            If chkMusi.value = 0 Then
                'sStr = sStr & "     AND BIGO1 <= (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB WHERE ACID = '" & Trim(basModule.SchCD) & "') "       '> 완료여부 : 저장되면 YYMM값이 들어감. : 2008.02.28
                sStr = sStr & "     AND BIGO1 <= (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB) "       '> 완료여부 : 저장되면 YYMM값이 들어감. : 2008.02.28
            End If
    End Select
    
    
    
    
    
    
    
    
'>> 완료여부 : 저장되면 YYMM값이 들어감.
    sStr = sStr & " AND CL_CLOSE IS NULL "
    sStr = sStr & " AND BIGO2 IS NULL"                      '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    'sStr = sStr & " AND sel2_sch ='E'"
        Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
    '>> 분원
        For ni = 1 To 4 Step 1
            sTmp = Trim(Right(cboHakwon.Text, 30))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        Next ni
        
    '>> 수험번호
        If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
            sTmp = Trim(fpExmID_S.UnFmtText)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            sTmp = Trim(fpExmID_E.UnFmtText)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
            sTmp = Trim(fpExmID_S.UnFmtText)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
            sTmp = Trim(fpExmID_S.UnFmtText)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("EXMID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
            ' no action
        End If
    
    '>> 학생명
'        If Trim(txtStdNM.Text) > " " Then
'            sTmp = "%" & Trim(txtStdNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    '>> 주민번호
'        If Trim(fpBirth_ymd.UnFmtText) > " " Then
'            sTmp = "%" & Trim(fpBirth_ymd.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    Text1.Text = sStr
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
    
        fpTotCnt.value = 0
    
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
            
                fpTotCnt.value = fpTotCnt.value + 1
                
                sprTamgu.MaxRows = sprTamgu.MaxRows + 1
                sprTamgu.Row = sprTamgu.MaxRows
                
                sprTamgu.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("GAEYUL_CD")) = False Then sTmp = Trim(.Fields("GAEYUL_CD"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        sKaeyol = sTmp
                        
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("CY_ACNT")) = False Then sTmp = Trim(.Fields("CY_ACNT"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 30, sTmp)
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("CY_ACNT2")) = False Then sTmp = Trim(.Fields("CY_ACNT2"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 30, sTmp)
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("CY_ACNT3")) = False Then sTmp = Trim(.Fields("CY_ACNT3"))
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 30, sTmp)
                sprTamgu.Col = sprTamgu.Col + 1
                    sTmp = " ":     If IsNull(.Fields("TOT_AMT")) = False Then nTmp = CDbl(.Fields("TOT_AMT"))
                        Call basFunction.Set_SprType_Numeric(sprTamgu, 0, 0, 999999999, ",", nTmp)
                                    
                sprTamgu.Col = sprTamgu.Col + 1:    Call basFunction.Set_SprType_ChkBox(sprTamgu)
                
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            
            On Error Resume Next
            '>> 선택과목 (사탐/ 과탐)
                For ni = 1 To SATAM_COUNT Step 1
                
                    If ni Mod 4 = 1 Then
                        sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                
                    sprTamgu.Col = sprTamgu.Col + 1
                    
                    Select Case ni
                        Case 1 To 8
                            sGbn = "SEL" & Trim(CStr(ni))
                        Case 9 To 11
                            If sKaeyol = "02" Or sKaeyol = "04" Or sKaeyol = "06" Then      '< 계열처리로 변경 : 2008.01.09
                                sGbn = "X"
                            Else
                                sGbn = "SEL" & Trim(CStr(ni))
                            End If
                    End Select
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        
                        If IsNull(.Fields(sGbn)) = False Then
                            If sTmp <> "" Then
                                Select Case sTmp
                                    Case constSatamCodes(0):  sTmp = constSatams(0)
                                    Case constSatamCodes(1):  sTmp = constSatams(1)
                                    Case constSatamCodes(2):  sTmp = constSatams(2)
                                    Case constSatamCodes(3):  sTmp = constSatams(3)
                                    Case constSatamCodes(4):  sTmp = constSatams(4)
                                    Case constSatamCodes(5):  sTmp = constSatams(5)
                                    Case constSatamCodes(6):  sTmp = constSatams(6)
                                    Case constSatamCodes(7):  sTmp = constSatams(7)
                                    Case constSatamCodes(8):  sTmp = constSatams(8)
                                    Case constSatamCodes(9):  sTmp = constSatams(9)
                                    'Case "11":  sTmp = "세지"
                                    
                                    Case "51":   sTmp = "물1"
                                    Case "52":   sTmp = "화1"
                                    Case "53":   sTmp = "생1"
                                    Case "54":   sTmp = "지1"
                                    Case "55":   sTmp = "물2"
                                    Case "56":   sTmp = "화2"
                                    Case "57":   sTmp = "생2"
                                    Case "58":   sTmp = "지2"
                                    
                                End Select
                            End If
                            Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                
                '사탐과목하나 줄면서 빈칸으로 처리
                sprTamgu.Col = sprTamgu.Col + 1
                Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprTamgu.Col = sprTamgu.Col + 1
                If IsNull(.Fields("SEL_X2")) = True Then
                    Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
                Else
                    If Trim(.Fields("SEL_X2")) = "00" Then
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
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
                            Case "42":  sTmp = "아랍어"
                            
                            Case "81":  sTmp = "미적분"
                            Case "82":  sTmp = "이산수학"
                            Case "83":  sTmp = "확률통계"
                            Case "84":  sTmp = "수리나형"
                            
                        End Select
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> 논술
                For ni = 1 To 4 Step 1
                    sprTamgu.Col = sprTamgu.Col + 1
                    
                    sGbn = "SEL_N" & Trim(CStr(ni))
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
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
                            Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
            '>> 금액
            '   2007.12.21 : 1 항목추가
            '   2008.01.09 : 1 항목 추가
                For ni = 1 To 10 Step 1
                    sprTamgu.Col = sprTamgu.Col + 1:    nTmp = 0
                    
                    If ni = 4 Then sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    If ni = 5 Then sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    If ni = 10 Then sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    
                    'sGbn = "BASE_AMT" & Trim(CStr(ni))
                    sGbn = "B" & Trim(CStr(ni))                 '< 2008.01.09
                    
                    If IsNull(.Fields(sGbn)) = False Then
                        nTmp = CDbl(.Fields(sGbn))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprTamgu, 0, 0, 999999999, ",", nTmp)
                Next ni
                
            '>> 탐구
            '   2007.12.21 : 1 항목추가
                For ni = 1 To 12 Step 1
                    sprTamgu.Col = sprTamgu.Col + 1:    nTmp = 0
                    
                    If ni Mod 4 = 0 Then
                        sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                    
                    sGbn = "TAMGU_AMT" & Trim(CStr(ni))
                    
                    If IsNull(.Fields(sGbn)) = False Then
                        nTmp = CDbl(.Fields(sGbn))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprTamgu, 0, 0, 999999999, ",", nTmp)
                Next ni
                
                
            '## formula ##      << 합계처리 : 열이 증가되면 반드시 수정해야 함.
                    sprTamgu.Col = 9
                    
                    sprTamgu.FormulaSync = False
                    sprTamgu.Formula = "SUM(X#:AV#)"            '< 추가 : 2007.12.21 : 2008.01.09
                
                .MoveNext
            Next nRec
            
            sprTamgu.Row = 1:       sprTamgu.Row2 = sprTamgu.MaxRows
            sprTamgu.Col = 1:       sprTamgu.Col2 = sprTamgu.MaxCols
            sprTamgu.BlockMode = True
                'sprTamgu.BackColor = basModule.BackColor2
                'sprTamgu.BackColorStyle = BackColorStyleUnderGrid
            sprTamgu.BlockMode = False

            sprTamgu.ColsFrozen = 8
            
        '>> spread lock
            sprTamgu.Row = 1:       sprTamgu.Row2 = sprTamgu.MaxRows
            sprTamgu.Col = 1:       sprTamgu.Col2 = 5
            sprTamgu.BlockMode = True
                sprTamgu.Lock = True
                sprTamgu.Protect = True
            sprTamgu.BlockMode = False
            
            sprTamgu.Row = 1:       sprTamgu.Row2 = sprTamgu.MaxRows
            sprTamgu.Col = 10:       sprTamgu.Col2 = 26
            sprTamgu.BlockMode = True
                sprTamgu.Lock = True
                sprTamgu.Protect = True
            sprTamgu.BlockMode = False
            
        End If
        
    End With
    
    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "합격생 조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "합격생 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "합격생 조회"
End Sub


'>> 선택 ## multi 선택
Private Sub sprTamgu_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    
    If Row < 1 Then Exit Sub

    With sprTamgu
        If .MaxRows < 1 Then Exit Sub

        'sprTamgu.Enabled = False
        
            If .Tag = "0" Then
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = &HFFFFFF
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = 10
                        .value = 0
                Next nRow
                
                .Row = Row:     .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = 10
                    .value = 1
                
                .Tag = Trim(CStr(Row))
            ElseIf .Tag > "0" Then
                .Row = Row
                .Col = 10
                If .value = 1 Then
                    .value = 0
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = &HFFFFFF
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    '.Tag = Trim(CStr(Row))
                Else
                    .value = 1
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    '.Tag = Trim(CStr(Row))
                End If
            
            End If
            
        'sprTamgu.Enabled = True

    End With
End Sub

Private Sub sprTamgu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nS      As Long
    Dim nE      As Long
    
    Dim nRow    As Long
    
    With sprTamgu
    
        If .MaxRows = 0 Then Exit Sub
        
        Select Case Shift
'            Case 0
'                Call sprTamgu_Click(1, .ActiveRow)
                
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
                            .Col = 10
                                .value = 1
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
    
    With sprTamgu
        If .MaxRows = 0 Then Exit Sub
            
        If chkAll.value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = 10
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = &HFFFFFF
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = 10
                    .value = 1
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
            
    End With
End Sub




'>> 금액등록
Private Sub cmdAmt_Click()
    Dim ni      As Long
    Dim nRec    As Long
    
    Dim sTmp    As String
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    With sprTamgu
    
    '## 등록전 금액정도 ini 파일에 등록 ---------------------------------------------------
    '> base
        Select Case Trim(basModule.SchCD)
            Case "N", "B"
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                    Case "02"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                        
                    '>> 노량진 계열별 처리
                    Case "03"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE03", sTmp)
                    Case "04"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE04", sTmp)
                    Case "05"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE05", sTmp)
                    Case "06"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE06", sTmp)
                        
                    Case "07"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE07", sTmp)
                    Case "08"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE08", sTmp)
                    Case "09"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE09", sTmp)
                    Case "10"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE10", sTmp)
                        
                        
                    Case "11"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE11", sTmp)
                    Case "12"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE12", sTmp)
                    Case "13"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE13", sTmp)
                    Case "14"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE14", sTmp)
                    Case "15"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE15", sTmp)
                    Case "16"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE16", sTmp)
                        
                    '## --------------------------------------------------##
                    
                End Select
                
            Case "K", "W", "Q", "M"                 '< 계열 : 2008.01.10 : 강남 2008.03.24
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01", "04", "06", "11", "16", "19"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                    Case "02", "05", "07", "12", "17", "20"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                End Select
            
            
            Case "S"
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                    Case "02"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                    
                    Case "03"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE03", sTmp)
                    
                    Case "05"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE05", sTmp)
                    Case "06"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE06", sTmp)
                    
                    Case "11"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE11", sTmp)
                    Case "12"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE12", sTmp)
                        
                    Case "18"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE11", sTmp)
                    Case "19"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE12", sTmp)
                    '## --------------------------------------------------##
                    
                End Select
                
            
            Case "J"
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                    Case "02"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                        
                    Case "11"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE11", sTmp)
                    Case "12"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE12", sTmp)
                        
                    Case "18"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE11", sTmp)
                    Case "19"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE12", sTmp)
                    '## --------------------------------------------------##
                    
                End Select
                
            Case Else
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01", "03", "11"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                    Case "02", "04", "12"                                           '< 계열처리 추가 : 2008.02.15
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                        
                        
                    Case "05"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE05", sTmp)
                    Case "06"
                        sTmp = ""
                        For ni = 1 To 10 Step 1
                            sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("BASE06", sTmp)
                        
                End Select
            
        End Select
    
    '> satam 사탐
        sTmp = ""
        For ni = 1 To SATAM_COUNT + 1 Step 1
            sTmp = sTmp & Trim(CStr(fpSatam(ni).value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("SATAM", sTmp)
    
    
    '> gwatam
        Select Case Trim(basModule.SchCD)
            Case "N"
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "04"       '< 수리나형
                        sTmp = ""
                        For ni = 1 To 9 Step 1
                            sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("GWATAMNA", sTmp)
                    Case Else
                        sTmp = ""
                        For ni = 1 To 9 Step 1
                            sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
                        Next ni
                        sTmp = insert_AMT_ini_File("GWATAM", sTmp)
                End Select
            Case Else
                sTmp = ""
                For ni = 1 To 9 Step 1
                    sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
                Next ni
                sTmp = insert_AMT_ini_File("GWATAM", sTmp)
        End Select
        
    '---------------------------------------------------------------------------------------
        
    
        If .MaxRows = 0 Then Exit Sub
        
        nRec = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 10                        '< 등록체크
            If .value = 1 Then
                nRec = 1
                Exit For
            End If
        Next ni
        
        If nRec = 0 Then
            MsgBox "금액 등록할 학생을 선택하여 주십시요.", vbExclamation + vbOKOnly, "금액 등록"
            Exit Sub
        End If
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 10
            
            If .value = 1 Then
            
                Select Case Trim(basModule.SchCD)           '< 2008.01.09
                    Case "N", "B"
                        Select Case Trim(Right(cboKaeyol.Text, 30))
                        
                        '>> 사탐
                            Case "01", "03", "05", "07", "09", "11", "13", "15"
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 31:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 32:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value        '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value        '< 변경
                                
                            '>> 사탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:  If StrComp(Trim(.Text), constSatams(0), vbTextCompare) = 0 Then .Col = 37:      .value = fpSatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), constSatams(1), vbTextCompare) = 0 Then .Col = 38:      .value = fpSatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), constSatams(2), vbTextCompare) = 0 Then .Col = 39:      .value = fpSatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), constSatams(3), vbTextCompare) = 0 Then .Col = 40:     .value = fpSatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), constSatams(4), vbTextCompare) = 0 Then .Col = 41:    .value = fpSatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), constSatams(5), vbTextCompare) = 0 Then .Col = 42:      .value = fpSatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), constSatams(6), vbTextCompare) = 0 Then .Col = 43:      .value = fpSatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), constSatams(7), vbTextCompare) = 0 Then .Col = 44:      .value = fpSatam(8).value
                                .Col = 19:  If StrComp(Trim(.Text), constSatams(8), vbTextCompare) = 0 Then .Col = 45:      .value = fpSatam(9).value
                                .Col = 20:  If StrComp(Trim(.Text), constSatams(9), vbTextCompare) = 0 Then .Col = 46:      .value = fpSatam(10).value
                                .Col = 21:  .value = ""
                                
                                .Col = 22:           '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "독어", "일어", "에스파냐", "에스파냐어", "에파", "중국", "중국어", "중어", "한문", "언어", "수리", "영어", "세계사", "세사", "세지", "아랍어"        '< 추가 : 2008.01.09
                                            .Col = 48:      .value = fpSatam(11).value
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                                    
                        '>> 과탐
                            Case "02", "04", "06", "08", "10", "12", "14", "16"
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 31:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 32:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value            '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value            '< 변경
                                
                            '>> 과탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:   If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Then .Col = 37:       .value = fpGwatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Then .Col = 38:       .value = fpGwatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Then .Col = 39:       .value = fpGwatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Then .Col = 40:       .value = fpGwatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Then .Col = 41:       .value = fpGwatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Then .Col = 42:       .value = fpGwatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Then .Col = 43:       .value = fpGwatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then .Col = 44:       .value = fpGwatam(8).value
                                
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                
                                .Col = 20:  '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "미적분"
                                            .Col = 48:      .value = fpGwatam(9).value
                                        Case "이산", "확률", "나형", "이산수학", "확률통계", "수리나형"
                                            If Trim(basModule.SchCD) = "N" Then         '< 노량진 요청사항
                                                .Col = 48:      .value = 0
                                            Else
                                                .Col = 48:      .value = fpGwatam(9).value
                                            End If
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                            
                        End Select

                    Case "S", "P", "J"           '< 송파/ 마송/양재 : 2008.02.15
                        Select Case Trim(Right(cboKaeyol.Text, 30))
                        
                        '>> 사탐
                            Case "01", "03", "05", "11", "18"
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 29:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 30:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value        '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value        '< 변경
                                
                            '>> 사탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:  If StrComp(Trim(.Text), constSatams(0), vbTextCompare) = 0 Then .Col = 37:      .value = fpSatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), constSatams(1), vbTextCompare) = 0 Then .Col = 38:      .value = fpSatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), constSatams(2), vbTextCompare) = 0 Then .Col = 39:      .value = fpSatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), constSatams(3), vbTextCompare) = 0 Then .Col = 40:      .value = fpSatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), constSatams(4), vbTextCompare) = 0 Then .Col = 41:      .value = fpSatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), constSatams(5), vbTextCompare) = 0 Then .Col = 42:      .value = fpSatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), constSatams(6), vbTextCompare) = 0 Then .Col = 43:      .value = fpSatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), constSatams(7), vbTextCompare) = 0 Then .Col = 44:      .value = fpSatam(8).value
                                .Col = 19:  If StrComp(Trim(.Text), constSatams(8), vbTextCompare) = 0 Then .Col = 45:      .value = fpSatam(9).value
                                .Col = 20:  If StrComp(Trim(.Text), constSatams(9), vbTextCompare) = 0 Then .Col = 46:      .value = fpSatam(10).value
                                .Col = 21:  .value = ""
                                
                                
                                .Col = 22:  '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "독어", "일어", "에스파냐", "에스파냐어", "에파", "중국", "중국어", "중어", "한문", "언어", "수리", "영어", "세계사", "세사", "세지", "아랍어"        '< 추가 : 2008.01.09
                                            .Col = 48:      .value = fpSatam(11).value
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                                    
                        '>> 과탐
                            Case "02", "04", "06", "08", "12", "19"
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 29:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 30:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value        '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value        '< 변경
                                
                            '>> 과탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:   If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Then .Col = 37:       .value = fpGwatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Then .Col = 38:       .value = fpGwatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Then .Col = 39:       .value = fpGwatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Then .Col = 40:       .value = fpGwatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Then .Col = 41:       .value = fpGwatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Then .Col = 42:       .value = fpGwatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Then .Col = 43:       .value = fpGwatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then .Col = 44:       .value = fpGwatam(8).value
                                
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                
                                .Col = 20:  '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "미적분"
                                            .Col = 48:      .value = fpGwatam(9).value
                                        Case "이산", "확률", "나형", "이산수학", "확률통계", "수리나형"
                                            If Trim(basModule.SchCD) = "N" Then             '< 노량진 요청사항
                                                .Col = 48:      .value = 0
                                            Else
                                                .Col = 48:      .value = fpGwatam(9).value
                                            End If
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                            
                        End Select

                    Case Else
                        Select Case Trim(Right(cboKaeyol.Text, 30))
                        
                        '>> 사탐
                            Case "01", "04", "06", "11", "16", "19"         '< 계열 : 2008.01.10 : 강남
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 31:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 32:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value            '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value            '< 변경
                                
                            '>> 사탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:  If StrComp(Trim(.Text), constSatams(0), vbTextCompare) = 0 Then .Col = 37:      .value = fpSatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), constSatams(1), vbTextCompare) = 0 Then .Col = 38:      .value = fpSatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), constSatams(2), vbTextCompare) = 0 Then .Col = 39:      .value = fpSatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), constSatams(3), vbTextCompare) = 0 Then .Col = 40:      .value = fpSatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), constSatams(4), vbTextCompare) = 0 Then .Col = 41:      .value = fpSatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), constSatams(5), vbTextCompare) = 0 Then .Col = 42:      .value = fpSatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), constSatams(6), vbTextCompare) = 0 Then .Col = 43:      .value = fpSatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), constSatams(7), vbTextCompare) = 0 Then .Col = 44:      .value = fpSatam(8).value
                                .Col = 19:  If StrComp(Trim(.Text), constSatams(8), vbTextCompare) = 0 Then .Col = 45:      .value = fpSatam(9).value
                                .Col = 20:  If StrComp(Trim(.Text), constSatams(9), vbTextCompare) = 0 Then .Col = 46:      .value = fpSatam(10).value
                                .Col = 21:  .value = ""
                                
                                
                                .Col = 22:  '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "독어", "일어", "에스파냐", "에스파냐어", "에파", "중국", "중국어", "중어", "한문", "언어", "수리", "영어", "세계사", "세사", "세지", "아랍어"         '< 추가 : 2008.01.09
                                            .Col = 48:      .value = fpSatam(11).value
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                                    
                        '>> 과탐
                            Case "02", "05", "07", "12", "17", "20"            '< 계열 : 2008.01.10 : 강남
                            '>> 기본금액
                                For nRec = 1 To 4 Step 1
                                    .Col = 27 + nRec - 1
                                    .value = fpBase(nRec).value
                                Next nRec
                                
                                .Col = 29:      .value = fpBase(9).value        '< 수익자 부담금 : 2007.12.21
                                .Col = 30:      .value = fpBase(10).value       '< 기타 : 2008.01.09
                                
                            '>> 논술금액
                                .Col = 33:  .value = 0
                                .Col = 34:  .value = 0
                                .Col = 35:  .value = 0
                                .Col = 36:  .value = 0
                                
                                .Col = 23:  If StrComp(Trim(.Text), "언어", vbTextCompare) = 0 Then .Col = 33:      .value = fpBase(5).value
                                .Col = 24:  If StrComp(Trim(.Text), "수리", vbTextCompare) = 0 Then .Col = 34:      .value = fpBase(6).value
                                .Col = 25:  If StrComp(Trim(.Text), "외국어", vbTextCompare) = 0 Then .Col = 35:    .value = fpBase(7).value        '< 변경
                                .Col = 26:  If StrComp(Trim(.Text), "", vbTextCompare) = 0 Then .Col = 36:          .value = fpBase(8).value        '< 변경
                                
                            '>> 과탐금액
                                .Col = 37:  .value = 0
                                .Col = 38:  .value = 0
                                .Col = 39:  .value = 0
                                .Col = 40:  .value = 0
                                .Col = 41:  .value = 0
                                .Col = 42:  .value = 0
                                .Col = 43:  .value = 0
                                .Col = 44:  .value = 0
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                .Col = 48:  .value = 0
                                
                                .Col = 11:   If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Then .Col = 37:       .value = fpGwatam(1).value
                                .Col = 12:  If StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Then .Col = 38:       .value = fpGwatam(2).value
                                .Col = 13:  If StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Then .Col = 39:       .value = fpGwatam(3).value
                                .Col = 14:  If StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Then .Col = 40:       .value = fpGwatam(4).value
                                .Col = 15:  If StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Then .Col = 41:       .value = fpGwatam(5).value
                                .Col = 16:  If StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Then .Col = 42:       .value = fpGwatam(6).value
                                .Col = 17:  If StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Then .Col = 43:       .value = fpGwatam(7).value
                                .Col = 18:  If StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then .Col = 44:       .value = fpGwatam(8).value
                                
                                .Col = 45:  .value = 0
                                .Col = 46:  .value = 0
                                .Col = 47:  .value = 0
                                
                                .Col = 20:  '< 제2 과목 : 2007.12.21
                                    Select Case Trim(.Text)
                                        Case "미적분"
                                            .Col = 48:      .value = fpGwatam(9).value
                                        Case "이산", "확률", "나형", "이산수학", "확률통계", "수리나형"
                                            If Trim(basModule.SchCD) = "N" Then             '< 노량진 요청사항
                                                .Col = 48:      .value = 0
                                            Else
                                                .Col = 48:     .value = fpGwatam(9).value
                                            End If
                                        Case Else
                                            .Col = 48:      .value = 0
                                    End Select
                            
                        End Select

                End Select
            End If
        Next ni
    End With
End Sub




'## SORT ORDER ##
Private Sub cmdSort_Click()
    Dim ni      As Integer
    Dim nj      As Integer
    Dim nC      As Integer
    
    Dim sSort   As String
    Dim sR      As String
    
    Dim sDivs() As String
    Dim sDivC() As String
    
    nC = 0
    sSort = ""
    
    With sprTamgu
        For ni = 1 To 3 Step 1
            For nj = 0 To 2 Step 1
                If fpSort(nj).value = ni Then
                    nC = nC + 1
                    
                    Select Case nj
                        Case 0                      '<< 수험번호
                            .SortKey(nC) = 2
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "0," & CInt(Trim(fpSort(0).value)) & "/"
                        Case 1                      '<< 성명
                            .SortKey(nC) = 3
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "1," & CInt(Trim(fpSort(1).value)) & "/"
                        Case 2                      '<< 계열
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "2," & CInt(Trim(fpSort(2).value)) & "/"
                    End Select
                    
                End If
            Next nj
        Next ni
        
        .Sort -1, -1, -1, -1, SortByRow
        
        sR = insert_AMT_ini_File("SORT", sSort)
        
        sDivs() = Split(sR, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
            
            fpSort(CInt(sDivC(0))).value = CInt(sDivC(1))
        Next ni
    
    End With
End Sub




'>> 학생 수강료 등록
Private Sub cmdSave_Click()
    Dim sTmp        As String
    Dim ni          As Integer
    
    Dim bRet        As Boolean
    Dim nCnt        As Long
    
    On Error GoTo ErrStmt
    
    '## 등록전 금액정도 ini 파일에 등록 ------------------------------------------------
    '> base
    Select Case Trim(basModule.SchCD)
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                Case "02"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                    
                '>> 계열별 추가
                Case "03"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE03", sTmp)
                Case "04"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE04", sTmp)
                Case "05"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE05", sTmp)
                Case "06"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE06", sTmp)
                    
                Case "07"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE07", sTmp)
                Case "08"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE08", sTmp)
                Case "09"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE09", sTmp)
                Case "10"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE10", sTmp)
                    
                Case "11"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE11", sTmp)
                Case "12"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE12", sTmp)
                Case "13"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE13", sTmp)
                Case "14"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE14", sTmp)
                Case "15"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE15", sTmp)
                Case "16"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE16", sTmp)
                    
                '## ----------------------------------------------------------- ##
                    
            End Select
        Case "K", "W", "Q", "M"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "04", "06", "11", "16"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                Case "02", "05", "07", "12", "17"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEJA", sTmp)
            End Select
        
        Case Else       '< 송파, 마송 계열 추가 됨.
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEIN", sTmp)
                Case "02", "04"                         '< 2008.02.15
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.01.09
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASEJA", sTmp)
                
                Case "05"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE05", sTmp)
                Case "06"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE06", sTmp)
                    
                    
                Case "11"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE11", sTmp)
                Case "12"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE12", sTmp)
                    
                Case "18"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE11", sTmp)
                Case "19"
                    sTmp = ""
                    For ni = 1 To 10 Step 1             '< 2008.02.01
                        sTmp = sTmp & Trim(CStr(fpBase(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("BASE12", sTmp)
                    
                    
                    
            End Select
        
    End Select
    
    
    '> satam 사탐
        sTmp = ""
        For ni = 1 To SATAM_COUNT + 1 Step 1
            sTmp = sTmp & Trim(CStr(fpSatam(ni).value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("SATAM", sTmp)
    
    '> gwatam
    Select Case Trim(basModule.SchCD)
        Case "N", "B"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "04"       '< 수리나형
                    sTmp = ""
                    For ni = 1 To 9 Step 1
                        sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("GWATAMNA", sTmp)
                Case Else
                    sTmp = ""
                    For ni = 1 To 9 Step 1
                        sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
                    Next ni
                    sTmp = insert_AMT_ini_File("GWATAM", sTmp)
            End Select
        Case Else
            sTmp = ""
            For ni = 1 To 9 Step 1
                sTmp = sTmp & Trim(CStr(fpGwatam(ni).value)) & "/"
            Next ni
            sTmp = insert_AMT_ini_File("GWATAM", sTmp)
    End Select
    
    
        
    '----------------------------------------------------------------------------------
    
     
    If sprTamgu.MaxRows = 0 Then Exit Sub
     
    bRet = False
    
    '>> 조건체크
    With sprTamgu
        If .MaxRows = 0 Then Exit Sub
        
        nCnt = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 10
            If .value = 1 Then
                .Col = 6
                If Trim(.Text) = "" Then
                    MsgBox "가상계좌가 없습니다.", vbExclamation + vbOKOnly, "등록금 등록"
                    Exit Sub
                End If
'                .Col = 7
'                If Trim(.Text) = "" Then
'                    MsgBox "가상계좌가 없습니다.", vbExclamation + vbOKOnly, "교재비 등록"
'                    Exit Sub
'                End If
                ' 조건없앰... 2007.12.21
'                .Col = 7
'                If .Value = 0 Then
'                    MsgBox "등록금이 0 입니다.", vbExclamation + vbOKOnly, "등록금 등록"
'                    Exit Sub
'                End If
                
                .Col = 9
                If Trim(.Text) = "" Then
                    MsgBox "금액이 없습니다.", vbExclamation + vbOKOnly, "등록금 등록"
                    Exit Sub
                End If
                
                nCnt = nCnt + 1
            End If
        Next ni
        
        If nCnt = 0 Then
            MsgBox "선택 1개 이상하십시요.", vbExclamation + vbOKOnly, "등록금 등록"
            Exit Sub
        End If
    End With
    
    cmdSave.Enabled = False
        bRet = Save_Amt_Data
        
    cmdSave.Enabled = True
    
    If bRet = True Then
        MsgBox "등록금 등록하였습니다.", vbInformation + vbOKOnly, "등록금 등록"
    Else
        MsgBox "등록금 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "등록금 등록"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "등록금 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "등록금 등록"
    On Error GoTo 0
End Sub

'저장
Private Function Save_Amt_Data() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nRow        As Long
    Dim sStr        As String
    Dim nExe        As Integer
    
    Dim nRec        As Long                                 '<< 처리해야 할 수
    Dim nTot        As Long                                 '<< 처리한 수
    
    bRet = False
    nRec = 0
    nTot = 0
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    For nRow = 1 To sprTamgu.MaxRows Step 1
        
        sprTamgu.Row = nRow
        sprTamgu.Col = 10
        
        If sprTamgu.value = 1 Then      '< 선택
        
            nRec = nRec + 1
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
        
            sStr = ""
            sStr = sStr & "  Update CLSTD01TB"
            sStr = sStr & "     SET CY_ACNT    = ?, "
            sStr = sStr & "         CY_ACNT2    = ?, "
            sStr = sStr & "         CY_ACNT3    = ?, "
            sStr = sStr & "         TOT_AMT    = ?, "
            
            sStr = sStr & "         BASE_AMT1  = ?, "
            sStr = sStr & "         BASE_AMT2  = ?, "
            sStr = sStr & "         BASE_AMT3  = ?, "
            sStr = sStr & "         BASE_AMT4  = ?, "
            
            sStr = sStr & "         BASE_AMT9  = ?, "           '< 추가 : 2007.12.21
            sStr = sStr & "         BASE_AMT10 = ?, "           '< 추가 : 2008.01.09
            
            sStr = sStr & "         BASE_AMT5  = ?, "
            sStr = sStr & "         BASE_AMT6  = ?, "
            sStr = sStr & "         BASE_AMT7  = ?, "
            sStr = sStr & "         BASE_AMT8  = ?, "
            
            sStr = sStr & "         TAMGU_AMT1 = ?, "
            sStr = sStr & "         TAMGU_AMT2 = ?, "
            sStr = sStr & "         TAMGU_AMT3 = ?, "
            sStr = sStr & "         TAMGU_AMT4 = ?, "
            sStr = sStr & "         TAMGU_AMT5 = ?, "
            sStr = sStr & "         TAMGU_AMT6 = ?, "
            sStr = sStr & "         TAMGU_AMT7 = ?, "
            sStr = sStr & "         TAMGU_AMT8 = ?, "
            sStr = sStr & "         TAMGU_AMT9 = ?, "
            sStr = sStr & "         TAMGU_AMT10= ?, "
            sStr = sStr & "         TAMGU_AMT11= ?, "
            sStr = sStr & "         TAMGU_AMT12= ? "            '< 추가 : 2007.12.21
            
            sStr = sStr & "   WHERE SCHNO = ? "
            'sStr = sStr & "     AND ACID  = ? "
        
            '>> 계좌번호
                sprTamgu.Col = 6
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("CY_ACNT", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                sprTamgu.Col = 7
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("CY_ACNT2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                sprTamgu.Col = 8
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("CY_ACNT3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 전체금액
                sprTamgu.Col = 9
                If Trim(sprTamgu.Text) = "" Then
                    nTmp = 0
                Else
                    nTmp = CLng(sprTamgu.value)
                End If
                    Set DBParam = DBCmd.CreateParameter("TOT_AMT", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
            '>> 등록금, 교재비, 부교재비, 종합회원비, 논술1 ~ 논술4
                For ni = 27 To 36 Step 1            '< 2008.01.09
                
                    Select Case ni
                        Case 27 To 30
                            sprTamgu.Col = ni
                            If Trim(sprTamgu.Text) = "" Then
                                nTmp = 0
                            Else
                                nTmp = CLng(sprTamgu.value)
                            End If
                                sTmp = "BASE_AMT" & Trim(CStr(ni - 26))         '< 변경 : 2007.12.21
                                Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

                        Case 31
                            sprTamgu.Col = ni
                            If Trim(sprTamgu.Text) = "" Then
                                nTmp = 0
                            Else
                                nTmp = CLng(sprTamgu.value)
                            End If
                                sTmp = "BASE_AMT9"                              '< 고정항목
                                Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                        
                        Case 32
                            sprTamgu.Col = ni
                            If Trim(sprTamgu.Text) = "" Then
                                nTmp = 0
                            Else
                                nTmp = CLng(sprTamgu.value)
                            End If
                                sTmp = "BASE_AMT10"                             '< 고정항목
                                Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                                
                        Case 33 To 36
                            sprTamgu.Col = ni
                            If Trim(sprTamgu.Text) = "" Then
                                nTmp = 0
                            Else
                                nTmp = CLng(sprTamgu.value)
                            End If
                                sTmp = "BASE_AMT" & Trim(CStr(ni - 28))         '< 변경 : 2007.12.21 : 2008.01.09
                                Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                        
                    End Select
                    
                Next ni
            
            '>> 선택금액 1 ~ 11 + 제2과목 : 추가 : 2007.12.21   : 2008.01.09
                For ni = 37 To 48 Step 1
                    nTmp = 0
                    
                    sprTamgu.Col = ni
                    If Trim(sprTamgu.Text) = "" Then
                        nTmp = 0
                    Else
                        nTmp = CLng(sprTamgu.value)
                    End If
                        sTmp = "TAMGU_AMT" & Trim(CStr(ni - 36))        '< 변경 : 2007.12.21 : 2008.01.09 : 1부터 ~
                        Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                Next ni
                
            
            '>> 학생코드
                sprTamgu.Col = 1
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'            '>> 학원코드 분류
'                sTmp = Trim(basModule.SchCD)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                    Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
            
            nTot = nTot + nExe
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
        
        End If
    Next nRow
    
    If nRec = nTot Then
        Save_Amt_Data = True
    Else
        Save_Amt_Data = False
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_Amt_Data = False
End Function

