VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form LSN100 
   Caption         =   "시간표 만들기 >> 반 구성하기"
   ClientHeight    =   11385
   ClientLeft      =   3420
   ClientTop       =   1770
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11385
   ScaleWidth      =   15510
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  '없음
      Caption         =   "Frame8"
      Height          =   1455
      Left            =   30
      TabIndex        =   40
      Top             =   9780
      Width           =   15465
      Begin FPSpread.vaSpread sprClassDet 
         Height          =   1095
         Left            =   90
         TabIndex        =   41
         Top             =   300
         Width           =   15315
         _Version        =   393216
         _ExtentX        =   27014
         _ExtentY        =   1931
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
         SpreadDesigner  =   "LSN100.frx":0000
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "반 처리 임시 Spread"
         Height          =   210
         Left            =   180
         TabIndex        =   42
         Top             =   60
         Width           =   2265
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame7"
      Height          =   3645
      Left            =   30
      TabIndex        =   37
      Top             =   6000
      Width           =   15435
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Caption         =   "Frame6"
         Height          =   3585
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   15375
         Begin VB.CommandButton cmdOrdGwamok_View 
            Caption         =   "학생신청과목 펼친내역 보기"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11910
            TabIndex        =   26
            Top             =   30
            Width           =   2655
         End
         Begin VB.CommandButton cmdProcClass 
            Caption         =   "반 설정하기 (반설정 조회하기 클릭 -> 반 처리할 학생을 선택 -> 설정하기 클릭하세요.)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   25
            Top             =   30
            Width           =   8505
         End
         Begin VB.CommandButton cmdDeleteClass 
            Caption         =   "선택 반 등록내역 삭제하기"
            Height          =   375
            Left            =   12720
            TabIndex        =   28
            Top             =   3180
            Width           =   2595
         End
         Begin VB.TextBox txtKaeyol 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2220
            TabIndex        =   24
            Text            =   "txtKaeyol"
            Top             =   75
            Width           =   675
         End
         Begin VB.CommandButton cmdClass 
            Caption         =   "반 설정조회"
            Height          =   375
            Left            =   60
            TabIndex        =   23
            Top             =   30
            Width           =   1575
         End
         Begin FPSpread.vaSpread sprClass 
            Height          =   2715
            Left            =   30
            TabIndex        =   39
            Top             =   450
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   4789
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
            SpreadDesigner  =   "LSN100.frx":01D4
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "스프레드 <delete 키 동작>"
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   60
            Top             =   3300
            Width           =   3405
         End
         Begin VB.Label Label6 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   1230
            TabIndex        =   45
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame5"
      Height          =   615
      Left            =   30
      TabIndex        =   34
      Top             =   30
      Width           =   15465
      Begin VB.Frame Frame4 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame4"
         Height          =   555
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   15405
         Begin VB.CommandButton cmdinput_Class 
            Caption         =   "반 등록하기"
            Height          =   495
            Left            =   13650
            TabIndex        =   43
            Top             =   30
            Width           =   1725
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '투명
            Caption         =   $"LSN100.frx":4621
            Height          =   375
            Left            =   5160
            TabIndex        =   46
            Top             =   90
            Width           =   8805
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   $"LSN100.frx":46B6
            Height          =   375
            Left            =   30
            TabIndex        =   44
            Top             =   90
            Width           =   6285
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   30
      TabIndex        =   29
      Top             =   660
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   5235
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   15375
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame6"
            Height          =   375
            Left            =   60
            TabIndex        =   53
            Top             =   450
            Width           =   8025
            Begin VB.ComboBox cboGwamok 
               Height          =   300
               Index           =   2
               Left            =   7020
               Style           =   2  '드롭다운 목록
               TabIndex        =   16
               Top             =   120
               Width           =   945
            End
            Begin VB.ComboBox cboGwamok 
               Height          =   300
               Index           =   1
               Left            =   6090
               Style           =   2  '드롭다운 목록
               TabIndex        =   15
               Top             =   120
               Width           =   945
            End
            Begin VB.ComboBox cboGwamok 
               Height          =   300
               Index           =   0
               Left            =   5160
               Style           =   2  '드롭다운 목록
               TabIndex        =   14
               Top             =   120
               Width           =   945
            End
            Begin VB.CommandButton cmdSort 
               BackColor       =   &H00C0C0FF&
               Caption         =   "정렬"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   22
               Top             =   30
               Width           =   645
            End
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   4
               Left            =   4500
               TabIndex        =   13
               Top             =   120
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
               MaxValue        =   "6"
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
               Left            =   2610
               TabIndex        =   10
               Top             =   120
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
               MaxValue        =   "6"
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
               Left            =   3240
               TabIndex        =   11
               Top             =   120
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
               MaxValue        =   "6"
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
               Index           =   3
               Left            =   3870
               TabIndex        =   12
               Top             =   120
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
               MaxValue        =   "6"
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
               Index           =   5
               Left            =   1890
               TabIndex        =   67
               Top             =   120
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
               MaxValue        =   "6"
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
               Index           =   6
               Left            =   1200
               TabIndex        =   69
               Top             =   150
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
               MaxValue        =   "5"
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
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "수험번호"
               Height          =   210
               Index           =   6
               Left            =   1080
               TabIndex        =   70
               Top             =   0
               Width           =   765
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "타입"
               Height          =   210
               Index           =   5
               Left            =   1890
               TabIndex        =   68
               Top             =   -30
               Width           =   465
            End
            Begin VB.Label Label18 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "과목"
               Height          =   210
               Left            =   7140
               TabIndex        =   63
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label14 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "과목"
               Height          =   210
               Left            =   6180
               TabIndex        =   62
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label13 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "과목"
               Height          =   210
               Left            =   5280
               TabIndex        =   61
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label19 
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
               Left            =   30
               TabIndex        =   58
               Top             =   105
               Width           =   645
            End
            Begin VB.Label Label7 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "합계"
               Height          =   210
               Left            =   4470
               TabIndex        =   57
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "언어"
               Height          =   210
               Index           =   1
               Left            =   2580
               TabIndex        =   56
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "수리"
               Height          =   210
               Index           =   2
               Left            =   3210
               TabIndex        =   55
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "외국어"
               Height          =   210
               Index           =   3
               Left            =   3810
               TabIndex        =   54
               Top             =   -15
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdDelStdClass 
            Caption         =   "선택학생 반 삭제하기"
            Height          =   315
            Left            =   12750
            TabIndex        =   27
            Top             =   4860
            Width           =   2595
         End
         Begin VB.CommandButton cmdNotProcDataSelect 
            Caption         =   "반 없음선택"
            Height          =   315
            Left            =   13920
            TabIndex        =   20
            Top             =   495
            Width           =   1245
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "선택"
            Height          =   315
            Left            =   14250
            TabIndex        =   21
            Top             =   840
            Width           =   885
         End
         Begin VB.CommandButton cmdFindStd 
            Caption         =   "학생 조회하기"
            Height          =   405
            Left            =   210
            TabIndex        =   0
            Top             =   30
            Width           =   1605
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   4050
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   105
            Width           =   915
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   2550
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   105
            Width           =   1035
         End
         Begin FPSpread.vaSpread sprSTD 
            Height          =   3975
            Left            =   60
            TabIndex        =   36
            Top             =   840
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   7011
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
            SpreadDesigner  =   "LSN100.frx":4733
         End
         Begin EditLib.fpMask fpExmID_S 
            Height          =   345
            Left            =   5790
            TabIndex        =   3
            Top             =   90
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            Mask            =   "AAAAA"
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
            Left            =   6900
            TabIndex        =   4
            Top             =   90
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            Mask            =   "AAAAA"
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
         Begin VB.Frame Frame9 
            BackColor       =   &H00D2EAF5&
            Height          =   495
            Left            =   8130
            TabIndex        =   47
            Top             =   -30
            Width           =   7275
            Begin EditLib.fpLongInteger fpTotS 
               Height          =   345
               Left            =   5430
               TabIndex        =   8
               Top             =   120
               Width           =   675
               _Version        =   196608
               _ExtentX        =   1191
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
            Begin EditLib.fpLongInteger fpTotE 
               Height          =   345
               Left            =   6510
               TabIndex        =   9
               Top             =   120
               Width           =   675
               _Version        =   196608
               _ExtentX        =   1191
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
            Begin EditLib.fpLongInteger fpKor 
               Height          =   345
               Left            =   450
               TabIndex        =   5
               Top             =   120
               Width           =   645
               _Version        =   196608
               _ExtentX        =   1138
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
            Begin EditLib.fpLongInteger fpMat 
               Height          =   345
               Left            =   2100
               TabIndex        =   6
               Top             =   120
               Width           =   645
               _Version        =   196608
               _ExtentX        =   1138
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
            Begin EditLib.fpLongInteger fpEng 
               Height          =   345
               Left            =   3810
               TabIndex        =   7
               Top             =   120
               Width           =   645
               _Version        =   196608
               _ExtentX        =   1138
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
            Begin VB.Label Label9 
               BackStyle       =   0  '투명
               Caption         =   "합계             부터"
               Height          =   210
               Left            =   5010
               TabIndex        =   51
               Top             =   180
               Width           =   1995
            End
            Begin VB.Label Label15 
               BackStyle       =   0  '투명
               Caption         =   "언어            이상/"
               Height          =   210
               Left            =   60
               TabIndex        =   50
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label16 
               BackStyle       =   0  '투명
               Caption         =   "수리            이상/"
               Height          =   210
               Left            =   1680
               TabIndex        =   49
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label17 
               BackStyle       =   0  '투명
               Caption         =   "외국어            이상/"
               Height          =   210
               Left            =   3240
               TabIndex        =   48
               Top             =   180
               Width           =   1755
            End
         End
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   8580
            TabIndex        =   64
            Top             =   4860
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00D2EAF5&
            Height          =   465
            Left            =   8130
            TabIndex        =   33
            Top             =   390
            Width           =   4605
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "전체학생"
               Height          =   315
               Index           =   2
               Left            =   3360
               TabIndex        =   19
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "반 설정된 학생"
               Height          =   315
               Index           =   1
               Left            =   1740
               TabIndex        =   18
               Top             =   120
               Width           =   1695
            End
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "반 미정인 학생"
               Height          =   315
               Index           =   0
               Left            =   90
               TabIndex        =   17
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '투명
            Caption         =   "스프레드 <delete 키 동작>"
            Height          =   210
            Index           =   4
            Left            =   9840
            TabIndex        =   66
            Top             =   4950
            Width           =   3405
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '투명
            Caption         =   "조회인원"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   7740
            TabIndex        =   65
            Top             =   4950
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '투명
            Caption         =   $"LSN100.frx":8B9D
            Height          =   360
            Left            =   150
            TabIndex        =   59
            Top             =   4860
            Width           =   11175
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '투명
            Caption         =   "수험번호            부터             까지"
            Height          =   210
            Left            =   5040
            TabIndex        =   52
            Top             =   150
            Width           =   3075
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   3030
            TabIndex        =   32
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시험형태"
            Height          =   210
            Left            =   1590
            TabIndex        =   31
            Top             =   150
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "LSN100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : LSN100
'   모 듈  목 적 : 시간표 만들기 >> 반 구성하기
'
'   작   성   일 : 2007/10/22
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Type tClass
    CLSCD   As String
    CLSNM   As String
End Type
Private Const nRowHeight = 14


Private Sub cmdOrdGwamok_View_Click()
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "반별 과목 신청내역 조회를 하십시요.", vbExclamation + vbOKOnly, "학생신청과목 펼친내역 보기"
        Exit Sub
    End If

    Load TMR022
    TMR022.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload TMR022
End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0, 15700, 9980
    
    Me.Tag = "LOAD"
        With sprSTD
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
        End With
        
        With sprClass
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
        End With
        
        With cboExmType
            .Clear
            .AddItem "전  체" & Space(30) & "ALL"
            .AddItem "무시험" & Space(30) & "0"
            .AddItem "유시험" & Space(30) & "1"
            .ListIndex = 0
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            '.AddItem "예체" & Space(30) & "03"
            
            .ListIndex = 0
            
            txtKaeyol.Text = Trim(cboKaeyol.Text)
        End With
        
        Call init_Form
        
    Me.Tag = ""
    
End Sub

'## 헤더변경
Private Sub cboKaeyol_Click()
    Dim sTmp        As String
    Dim ni          As Integer
    
    txtKaeyol.Text = Trim(cboKaeyol.Text)
    
    With sprSTD
        Select Case Trim(Right(cboKaeyol.Text, 30))
            Case "01", "03"         '<< 인문
                
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                '.MaxCols = 21
                .MaxCols = 26           '< 2007.12.17
                
                .Col = 1:           .Text = "학생":         .ColWidth(.Col) = 7.2
                .Col = .Col + 1:    .Text = "학생명":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "수험":         .ColWidth(.Col) = 5
                
                '< 2007.12.17 ------------------------------------------------------
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "외국":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "합계":         .ColWidth(.Col) = 4
                '-------------------------------------------------------------------
                
                .Col = .Col + 1:    .Text = "국사":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "윤리":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "경제":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "한근":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "세계사":       .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "경지":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "한지":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "정치":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "사문":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "법사":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "세지":         .ColWidth(.Col) = 4.5
                
                .Col = .Col + 1:    .Text = "제2외":        .ColWidth(.Col) = 4.5
               
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "사탐":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "과탐":         .ColWidth(.Col) = 4.5
                
                .Col = .Col + 1:    .Text = "반넣기":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "시험종류":     .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "선택":         .ColWidth(.Col) = 6
                
            Case "02"       '<< 자연
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                '.MaxCols = 18
                .MaxCols = 23           '< 2007.12.17
                
                .Col = 1:           .Text = "학생":         .ColWidth(.Col) = 7.2
                .Col = .Col + 1:    .Text = "학생명":       .ColWidth(.Col) = 7
                .Col = .Col + 1:    .Text = "수험":         .ColWidth(.Col) = 6
                
                '< 2007.12.17 ------------------------------------------------------
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "외국":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "합계":         .ColWidth(.Col) = 5
                '-------------------------------------------------------------------
                
                .Col = .Col + 1:    .Text = "물1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "화1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "생1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "지1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "물2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "화2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "생2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "지2":          .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "사탐":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "과탐":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "반넣기":       .ColWidth(.Col) = 7.3
                .Col = .Col + 1:    .Text = "시험종류":     .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "선택":         .ColWidth(.Col) = 5.9
                
        End Select
        
        .MaxRows = 0
    End With
    
    
    With sprClass
        Select Case Trim(Right(cboKaeyol.Text, 30))
            Case "01", "03"         '<< 인문
                
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 21
                
                .Col = 1:           .Text = "반":           .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "반명":         .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "총원":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "선택":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "남은인원":     .ColWidth(.Col) = 8
                
                .Col = .Col + 1:    .Text = "국사":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "윤리":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "경제":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "한근":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "세계사":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "경지":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "한지":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "정치":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "사문":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "법사":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "세지":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "제2외":        .ColWidth(.Col) = 5
               
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "사탐":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "과탐":         .ColWidth(.Col) = 5
                
            Case "02"       '<< 자연
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 18
                
                .Col = 1:           .Text = "반":           .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "반명":         .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "총원":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "선택":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "남은인원":     .ColWidth(.Col) = 8
                
                .Col = .Col + 1:    .Text = "물1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "화1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "생1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "지1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "물2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "화2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "생2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "지2":          .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "언어":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "수리":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "사탐":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "과탐":         .ColWidth(.Col) = 5
                
        End Select
        
        .MaxRows = 0
    End With
    
    
    For ni = 0 To 2 Step 1
        With cboGwamok(ni)
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"         '<< 인문
                    .AddItem "없음" & Space(30) & "X"
                    .AddItem "국사" & Space(30) & "8"
                    .AddItem "윤리" & Space(30) & "9"
                    .AddItem "경제" & Space(30) & "10"
                    .AddItem "한근" & Space(30) & "11"
                    .AddItem "세계사" & Space(30) & "12"
                    .AddItem "경지" & Space(30) & "13"
                    .AddItem "한지" & Space(30) & "14"
                    .AddItem "정치" & Space(30) & "15"
                    .AddItem "사문" & Space(30) & "16"
                    .AddItem "법사" & Space(30) & "17"
                    .AddItem "세지" & Space(30) & "18"
                                         
'                    .AddItem "제2외" & Space(30) & "19"
'
'                    .AddItem "언어" & Space(30) & "20"
'                    .AddItem "수리" & Space(30) & "21"
'                    .AddItem "사탐" & Space(30) & "22"
'                    .AddItem "과탐" & Space(30) & "23"
                Case "02"
                    .AddItem "없음" & Space(30) & "X"
                    .AddItem "물1" & Space(30) & "8"
                    .AddItem "화1" & Space(30) & "9"
                    .AddItem "생1" & Space(30) & "10"
                    .AddItem "지1" & Space(30) & "11"
                    .AddItem "물2" & Space(30) & "12"
                    .AddItem "화2" & Space(30) & "13"
                    .AddItem "생2" & Space(30) & "14"
                    .AddItem "지2" & Space(30) & "15"
                    
                    
            End Select
            
            .ListIndex = 0
            
        End With
    Next ni
    
End Sub



Private Sub init_Form()
    
    fpTotCnt.value = 0
    
    optClass(0).value = True
    optClass(1).value = False
    optClass(2).value = False
    
    fpExmID_S.Text = ""
    fpExmID_E.Text = ""
    
    If Me.Tag = "LOAD" Then Exit Sub

        cboKaeyol.ListIndex = 0
    
End Sub















'>> 학생 조회
Private Sub cmdFindStd_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nCls        As Integer
    
    Dim sGbn        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    sprSTD.MaxRows = 0
    chkAll.value = 0
    fpTotCnt.value = 0
    
    sStr = ""
    sStr = sStr & "  SELECT A.SCHNO,"
    'sStr = sStr & "         GET_STDNM(A.SCHNO, A.ACID) AS STDNM, "
    sStr = sStr & "         B.STDNM, "
    sStr = sStr & "         B.EXMID AS GWANRIBUNHO ,"
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'01|') > 0 THEN          /* 사탐-국사 */"
    sStr = sStr & "             '01'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'02|') > 0 THEN          /* 사탐-윤리 */"
    sStr = sStr & "             '02'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'03|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "             '03'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'53|') > 0 THEN     /* 과탐-생물1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'04|') > 0 THEN          /* 사탐-한국근현대 */"
    sStr = sStr & "             '04'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         END END SEL4,"
    
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'05|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "             '05'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'06|') > 0 THEN          /* 사탐-경제지리 */"
    sStr = sStr & "             '06'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'07|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "             '07'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'57|') > 0 THEN     /* 과탐-생물2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'08|') > 0 THEN          /* 사탐-정치 */"
    sStr = sStr & "             '08'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         END END SEL8,"
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"       '<< 인문
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'09|') > 0 THEN          /* 사탐-사회문화 */"
            sStr = sStr & "             '09'"
            sStr = sStr & "         END SEL9,"
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'10|') > 0 THEN          /* 사탐-법과사회 */"
            sStr = sStr & "             '10'"
            sStr = sStr & "         END SEL10,"
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'11|') > 0 THEN          /* 사탐-세계지리 */"
            sStr = sStr & "             '11'"
            sStr = sStr & "         END SEL11,"
    End Select
    
    sStr = sStr & "  "
    sStr = sStr & "      /* 제2외국어 & 수리 */"
    sStr = sStr & "              CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'31|') > 0 THEN '31'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'32|') > 0 THEN '32'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'33|') > 0 THEN '33'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'34|') > 0 THEN '34'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'35|') > 0 THEN '35'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'36|') > 0 THEN '36'"
    
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'37|') > 0 THEN '37'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'38|') > 0 THEN '38'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'39|') > 0 THEN '39'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'40|') > 0 THEN '40'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'41|') > 0 THEN '41'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL2,'42|') > 0 THEN '42'"
    
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL4,'81|') > 0 THEN '81'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL4,'82|') > 0 THEN '82'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL4,'83|') > 0 THEN '83'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL4,'84|') > 0 THEN '84'"
    sStr = sStr & "         END END END END END END END END END END END END END END END END SEL_X2,"
    
    sStr = sStr & "      /* 논술 */"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'91|') > 0 THEN         /* 언어 */"
    sStr = sStr & "             '91'"
    sStr = sStr & "         END SEL_N1,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'92|') > 0 THEN         /* 수리 */"
    sStr = sStr & "             '92'"
    sStr = sStr & "         END SEL_N2,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'93|') > 0 THEN         /* 사탐 */"
    sStr = sStr & "             '93'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'94|') > 0 THEN         /* 과탐 */"
    sStr = sStr & "             '94'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         GET_LSNNM(A.ACID, A.SEL_CLASS) AS LSNNM, "
    
    '< 2007.12.17 ----------------------------------------------------------------------------------------
    sStr = sStr & "         A.K_NUM, A.M_NUM, A.E_NUM, A.TOT_NUM "
    sStr = sStr & "         , DECODE(B.MU_TYPE,'1','1수능','2','6월','3','9월','4','6월','5','9월') AS MU_TYPE "
    '-----------------------------------------------------------------------------------------------------
    
    sStr = sStr & "    FROM CLTTL01TB A, CLSTD01TB B"
    sStr = sStr & "   WHERE A.SCHNO = B.SCHNO "
    
    sStr = sStr & "     AND A.ACID  = B.ACID  "
    
    sStr = sStr & "     AND A.SCHNO > ' ' "
    
    sStr = sStr & "     AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    
'>> 계열
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "XX"
            ' no action
        Case "01", "03"
            sStr = sStr & " AND A.SEL1 > ' ' "
        Case "02"
            sStr = sStr & " AND A.SEL3 > ' ' "
    End Select
    If optClass(0).value = True Then
        sStr = sStr & " AND A.SEL_CLASS IS NULL "
    ElseIf optClass(1).value = True Then
        sStr = sStr & " AND A.SEL_CLASS > ' ' "
    ElseIf optClass(2).value = True Then
        ' no action
    End If
'>> 시험구분 (EXMTYPE)
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "ALL"
        
        Case "0"
            sStr = sStr & " AND A.EXMTYPE = '0' "
        Case "1"
            sStr = sStr & " AND A.EXMTYPE = '1' "
    End Select
'>> 수험번호
    'sStr = sStr & "     AND B.EXMID BETWEEN '" & Format(fpGwanri1.Value, "00000") & "'"
    'sStr = sStr & "                     AND '" & Format(fpGwanri2.Value, "00000") & "'"
'>> 수험번호            2007.12.17
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
    
'>> 합계
    If fpTotS.value > 0 And fpTotE.value > 0 Then
        sStr = sStr & " AND ( B.TOT_NUM >= " & Trim(CStr(fpTotS.value)) & " AND B.TOT_NUM <= " & Trim(CStr(fpTotE.value)) & ")"
    ElseIf fpTotS.value > 0 And fpTotE.value = 0 Then
        sStr = sStr & " AND ( B.TOT_NUM >= 0 AND B.TOT_NUM <= " & Trim(CStr(fpTotE.value)) & ")"
    ElseIf fpTotS.value = 0 And fpTotE.value > 0 Then
        sStr = sStr & " AND ( B.TOT_NUM >= " & Trim(CStr(fpTotS.value)) & " AND B.TOT_NUM <= 9999 )"
    Else
        ' no action
    End If
    
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "0"        '< 무시험
        '>> 언어
            If fpKor.value > 0 Then
                sStr = sStr & " AND B.K_NUM <= " & Trim(CStr(fpKor.value))
            End If
        '>> 수리
            If fpMat.value > 0 Then
                sStr = sStr & " AND B.M_NUM <= " & Trim(CStr(fpMat.value))
            End If
        '>> 외국어
            If fpEng.value > 0 Then
                sStr = sStr & " AND B.E_NUM <= " & Trim(CStr(fpEng.value))
            End If
        Case "1"        '< 유시험
        '>> 언어
            If fpKor.value > 0 Then
                sStr = sStr & " AND B.K_NUM >= " & Trim(CStr(fpKor.value))
            End If
        '>> 수리
            If fpMat.value > 0 Then
                sStr = sStr & " AND B.M_NUM >= " & Trim(CStr(fpMat.value))
            End If
        '>> 외국어
            If fpEng.value > 0 Then
                sStr = sStr & " AND B.E_NUM >= " & Trim(CStr(fpEng.value))
            End If
    End Select
    
'>> 완료여부 : 저장되면 YYMM값이 들어감.
    sStr = sStr & "     AND A.CL_CLOSE IS NULL "
    
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "     AND BIGO1 > 17"                     '< 2009.01.
    Else
        sStr = sStr & "     AND BIGO2 IS NULL"                  '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    End If
    
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
 
        
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

'   >> 관리번호
'        sTmp = Format(fpGwanri1.Value, "00000")
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
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
            
                fpTotCnt.value = fpTotCnt.value + 1
                
                sprSTD.MaxRows = sprSTD.MaxRows + 1
                sprSTD.Row = sprSTD.MaxRows:    sprSTD.RowHeight(sprSTD.Row) = nRowHeight
                
                sprSTD.Col = 1
                    sTmp = " ": If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("GWANRIBUNHO")) = False Then sTmp = Trim(.Fields("GWANRIBUNHO"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                    
            '> 2007.12.17 ----------------------------------------------------------------------------------------------------------------------
                sprSTD.Col = sprSTD.Col + 1     ' 국어
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = CLng(.Fields("K_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' 수학
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = CLng(.Fields("M_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' 영어
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = CLng(.Fields("E_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' 합계
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = CLng(.Fields("TOT_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                    
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            '-----------------------------------------------------------------------------------------------------------------------------------
            
                
            '>> 선택과목 (사탐/ 과탐)
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01", "03"
                        nCls = 11
                    Case "02"
                        nCls = 8
                End Select
            
                For ni = 1 To nCls Step 1
                
                    If ni Mod 4 = 1 And sprSTD.Col > 3 Then
                        sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                
                    sprSTD.Col = sprSTD.Col + 1
                    
                    Select Case ni
                        Case 1 To 8
                            sGbn = "SEL" & Trim(CStr(ni))
                        Case 9 To 11
                            If sKaeyol = "02" Then
                                sGbn = "X"
                            Else
                                sGbn = "SEL" & Trim(CStr(ni))
                            End If
                    End Select
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = " "
                        If IsNull(.Fields(sGbn)) = False Then
                            sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        End If
                                                
                        If sTmp <> "" Then
                            Select Case sTmp
                                Case "01":  sTmp = "국사"
                                Case "02":  sTmp = "윤리"
                                Case "03":  sTmp = "경제"
                                Case "04":  sTmp = "한근"
                                Case "05":  sTmp = "세계사"
                                Case "06":  sTmp = "경지"
                                Case "07":  sTmp = "한지"
                                Case "08":  sTmp = "정치"
                                Case "09":  sTmp = "사문"
                                Case "10":  sTmp = "법사"
                                Case "11":  sTmp = "세지"

                                Case "51":   sTmp = "물1"
                                Case "52":   sTmp = "화1"
                                Case "53":   sTmp = "생1"
                                Case "54":   sTmp = "지1"
                                Case "55":   sTmp = "물2"
                                Case "56":   sTmp = "화2"
                                Case "57":   sTmp = "생2"
                                Case "58":   sTmp = "지2"

                            End Select
                            
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                        
                    End If
                Next ni
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprSTD.Col = sprSTD.Col + 1
                If IsNull(.Fields("SEL_X2")) = True Then
                    Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 10, "")
                Else
                    sTmp = " "
                    If Trim(.Fields("SEL_X2")) = "00" Then
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 10, "")
                    Else
                        Select Case Trim(.Fields("SEL_X2"))

                            Case "31":  sTmp = "독어"
                            Case "32":  sTmp = "일어"
                            Case "33":  sTmp = "에스파냐어"
                            Case "34":  sTmp = "불어"
                            Case "35":  sTmp = "중국어"
                            Case "36":  sTmp = "한문"
                            
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
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> 논술
                For ni = 1 To 4 Step 1
                    sprSTD.Col = sprSTD.Col + 1
                    
                    sGbn = "SEL_N" & Trim(CStr(ni))
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = " "
                        If IsNull(.Fields(sGbn)) = False Then
                            sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        End If
                    
                        If sTmp <> "" Then
                            Select Case sTmp
                                Case "91":  sTmp = "언어"
                                Case "92":  sTmp = "수리"
                                Case "93":  sTmp = "외국어"     '< 변경
                                Case "94":  sTmp = ""

                            End Select
                            
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                            
                    End If
                Next ni
                
                sprSTD.Col = sprSTD.Col + 1
                    If IsNull(.Fields("LSNNM")) = True Then
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 50, "")
                    Else
                        sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                    
                
                sprSTD.Col = sprSTD.Col + 1
                    If IsNull(.Fields("MU_TYPE")) = True Then
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", 50, "")
                    Else
                        sTmp = Trim(.Fields("MU_TYPE"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                
                
                sprSTD.Col = sprSTD.MaxCols
                    Call basFunction.Set_SprType_ChkBox(sprSTD)
                    sprSTD.value = 0
                    
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                .MoveNext
            Next nRec
            
            sprSTD.Row = 1:       sprSTD.Row2 = sprSTD.MaxRows
            sprSTD.Col = 1:       sprSTD.Col2 = sprSTD.MaxCols
            sprSTD.BlockMode = True
                sprSTD.BackColor = basModule.WhiteColor
                sprSTD.BackColorStyle = BackColorStyleUnderGrid
            sprSTD.BlockMode = False

            sprSTD.ColsFrozen = 2
            
        '>> spread lock
        sprSTD.Row = 1:       sprSTD.Row2 = sprSTD.MaxRows
        sprSTD.Col = 1:       sprSTD.Col2 = sprSTD.MaxCols - 3
        sprSTD.BlockMode = True
            sprSTD.Lock = True
            sprSTD.Protect = True
        sprSTD.BlockMode = False
        
        sprSTD.Row = 1:                 sprSTD.Row2 = sprSTD.MaxRows
        sprSTD.Col = sprSTD.MaxCols:    sprSTD.Col2 = sprSTD.MaxCols
        sprSTD.BlockMode = True
            sprSTD.Lock = True
            sprSTD.Protect = True
        sprSTD.BlockMode = False
        
        
        End If
    End With
    
    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "학생조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "반 등록할 학생 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"

End Sub



Private Sub sprSTD_KeyUp(KeyCode As Integer, Shift As Integer)

    If sprSTD.ActiveRow < 1 Then Exit Sub

    Select Case KeyCode
        Case vbKeyDelete
            With sprSTD
                .Row = .ActiveRow
                .DeleteRows .Row, 1
                .MaxRows = .MaxRows - 1
            End With
    End Select
End Sub


Private Sub sprSTD_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow    As Long

    If Row = 0 Then
        Call cmdSort_Click
    End If
    
    If Row < 1 Then Exit Sub
    
    With sprSTD
        .Tag = "0"
        
        'no action
    End With
End Sub



Private Sub sprSTD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nRow        As Long
    
    With sprSTD
    
        If .ActiveRow < 1 Then Exit Sub
    
        Select Case Shift
            
            Case 0      'shift
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = .MaxCols
                        .value = 0
                Next nRow
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
        
                .Row = .ActiveRow: .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
        
                .Col = .MaxCols:    .value = 1

            Case 2      'ctrl
                .Row = .ActiveRow
                .Col = .MaxCols
                If .value = 1 Then
                    .Row = .ActiveRow:  .Row2 = .ActiveRow
                    .Col = 1:           .Col2 = .MaxCols
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Col = .MaxCols:    .value = 0
                    
                Else
                    .Row = .ActiveRow:  .Row2 = .ActiveRow
                    .Col = 1:           .Col2 = .MaxCols
                    .BlockMode = True
                        .BackColor = basModule.SelectColor2
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Col = .MaxCols:    .value = 1
                    
                End If
                
            Case 4
                .Tag = "4"
                
            Case Else
                
        End Select
        
    End With
End Sub

Private Sub sprSTD_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim nR1     As Long
    Dim nR2     As Long
    Dim nT      As Long
    Dim nRow    As Long
    
    If BlockRow < 1 Then
        BlockRow = 1
    End If
    If BlockRow2 < 1 Then
        BlockRow = 1
    End If
    
    
    nR1 = BlockRow
    nR2 = BlockRow2
    If BlockRow > BlockRow2 Then
        nT = nR1
        nR1 = nR2
        nR2 = nT
    End If
    
    With sprSTD
        Select Case .Tag
            Case "0"
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = .MaxCols
                        .value = 0
                Next nRow
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                For nRow = nR1 To nR2 Step 1
                    .Row = nRow
                    .Col = .MaxCols
                        .value = 1
                Next nRow
                
                .Row = nR1:     .Row2 = nR2
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
            Case "2"
                            
        End Select
    End With
End Sub


Private Sub chkAll_Click()
    Dim nRow        As Long
    
    With sprSTD
        Select Case chkAll.value
            Case 1
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = .MaxCols
                        .value = 1
                Next nRow
                
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            Case 0
                For nRow = 1 To .MaxRows Step 1
                    .Row = nRow
                    .Col = .MaxCols
                        .value = 0
                Next nRow
                
                .Row = 1:   .Row2 = .MaxRows
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
        End Select
    End With
    
End Sub



'>> 반 선택되지 않은 내용 row 선택
Private Sub cmdNotProcDataSelect_Click()
    Dim nRow        As Long
    
    With sprSTD
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2
            If Trim(.Text) = "" Then
                .Col = .MaxCols
                    .value = 1
                
                .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            Else
                .Col = .MaxCols
                    .value = 0
                
                .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End If
        Next nRow
    End With
    
End Sub








'<< 계열선택시 계열에 해당하는 반조회
Private Sub cmdClass_Click()
    If Trim(txtKaeyol.Text) = "" Then
        Exit Sub
    End If
    
    Call Find_Kaeyol_to_Class(Trim(Right(txtKaeyol.Text, 30)))
    
End Sub

Private Sub Find_Kaeyol_to_Class(ByVal aKaeyol As String)
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nCls        As Integer
    
    Dim sGbn        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    sprClass.MaxRows = 0
    
    sStr = ""
    sStr = sStr & " SELECT ACID, LSNCD, LSNNM, "
    sStr = sStr & "        LSNCAPA, SEL_OK, PROC_NO, "
    sStr = sStr & "        SEL1, SEL2, SEL3, SEL4, SEL5, SEL6, SEL7, SEL8,"
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"       '<< 인문
            sStr = sStr & " SEL9, SEL10, SEL11, "
    End Select
    sStr = sStr & "        SEL_X2, "
    sStr = sStr & "        SEL_N1, SEL_N2, SEL_N3, SEL_N4 "
    sStr = sStr & "   FROM ("
    
            sStr = sStr & " SELECT A.ACID, LSNCD, MAX(LSNNM) AS LSNNM, MAX(LSNCDNM) AS LSNCDNM, "
            sStr = sStr & "        MAX(LSNCAPA) AS LSNCAPA, COUNT(SEL_OK) AS SEL_OK, MAX(LSNCAPA)-COUNT(SEL_OK) /*MAX(PROC_NO)*/ AS PROC_NO,"
            
            sStr = sStr & "        SUM(NVL(SEL1 ,0))    AS SEL1 ,"
            sStr = sStr & "        SUM(NVL(SEL2 ,0))    AS SEL2 ,"
            sStr = sStr & "        SUM(NVL(SEL3 ,0))    AS SEL3 ,"
            sStr = sStr & "        SUM(NVL(SEL4 ,0))    AS SEL4 ,"
            sStr = sStr & "        SUM(NVL(SEL5 ,0))    AS SEL5 ,"
            sStr = sStr & "        SUM(NVL(SEL6 ,0))    AS SEL6 ,"
            sStr = sStr & "        SUM(NVL(SEL7 ,0))    AS SEL7 ,"
            sStr = sStr & "        SUM(NVL(SEL8 ,0))    AS SEL8 ,"
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"       '<< 인문
                    sStr = sStr & "        SUM(NVL(SEL9 ,0))    AS SEL9 ,"
                    sStr = sStr & "        SUM(NVL(SEL10,0))    AS SEL10,"
                    sStr = sStr & "        SUM(NVL(SEL11,0))    AS SEL11,"
            
            End Select
            sStr = sStr & "        SUM(NVL(SEL_X2,0))   AS SEL_X2,"
            
            sStr = sStr & "        SUM(NVL(SEL_N1,0))   AS SEL_N1,"
            sStr = sStr & "        SUM(NVL(SEL_N2,0))   AS SEL_N2,"
            sStr = sStr & "        SUM(NVL(SEL_N3,0))   AS SEL_N3,"
            sStr = sStr & "        SUM(NVL(SEL_N4,0))   AS SEL_N4 "
            
            sStr = sStr & "   FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, LSNCAPA, SEL_OK /*, (LSNCAPA-SEL_OK) AS PROC_NO*/"
            sStr = sStr & "           FROM SDLSN01TB"
            sStr = sStr & "          WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "            AND KAEYOL  = '" & aKaeyol & "'"
            sStr = sStr & "         ) A,"
            sStr = sStr & "        (SELECT ACID, SEL_CLASS, SCHNO,"
            sStr = sStr & "                NVL(SEL1 ,0)     AS SEL1 ,"
            sStr = sStr & "                NVL(SEL2 ,0)     AS SEL2 ,"
            sStr = sStr & "                NVL(SEL3 ,0)     AS SEL3 ,"
            sStr = sStr & "                NVL(SEL4 ,0)     AS SEL4 ,"
            sStr = sStr & "                NVL(SEL5 ,0)     AS SEL5 ,"
            sStr = sStr & "                NVL(SEL6 ,0)     AS SEL6 ,"
            sStr = sStr & "                NVL(SEL7 ,0)     AS SEL7 ,"
            sStr = sStr & "                NVL(SEL8 ,0)     AS SEL8 ,"
            
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"       '<< 인문
                    sStr = sStr & "                NVL(SEL9 ,0)     AS SEL9 ,"
                    sStr = sStr & "                NVL(SEL10,0)     AS SEL10,"
                    sStr = sStr & "                NVL(SEL11,0)     AS SEL11,"
            End Select
            
            sStr = sStr & "                NVL(SEL_X2,0)    AS SEL_X2,"
            
            sStr = sStr & "                NVL(SEL_N1,0)    AS SEL_N1 ,"
            sStr = sStr & "                NVL(SEL_N2,0)    AS SEL_N2 ,"
            sStr = sStr & "                NVL(SEL_N3,0)    AS SEL_N3 ,"
            sStr = sStr & "                NVL(SEL_N4,0)    AS SEL_N4 "
            
            sStr = sStr & "           FROM (SELECT ACID, SEL_CLASS, SCHNO, STDNM,"
            
            sStr = sStr & "                    /* 사탐, 과탐 분리 */"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* 사탐-국사 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL1,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* 사탐-윤리 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL2,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* 사탐-경제 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생물1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL3,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* 사탐-한국근현대 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL4,"
            
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* 사탐-세계사 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL5,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* 사탐-경제지리 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL6,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* 사탐-한국지리 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생물2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL7,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* 사탐-정치 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL8,"
            
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"       '<< 인문
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* 사탐-사회문화 */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL9,"
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* 사탐-법과사회 */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL10,"
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* 사탐-세계지리 */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL11,"
            End Select
            
            sStr = sStr & "                 /* 제2외국어 & 수리 */"
            sStr = sStr & "                             CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN 1"
            
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN 1"
            
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN 1"
            sStr = sStr & "                        ELSE 0"
            sStr = sStr & "                        END END END END END END END END END END END END END END END END SEL_X2,"
            
            sStr = sStr & "                 /* 논술 */"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N1,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N2,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"       '< 변경
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N3,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"             '< 변경
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N4 "
            
            sStr = sStr & "                   FROM CLTTL01TB"
            sStr = sStr & "                  WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "                    AND SEL_CLASS"
            sStr = sStr & "                     IN (SELECT LSNCD"
            sStr = sStr & "                           FROM SDLSN01TB"
            sStr = sStr & "                          WHERE KAEYOL  = '" & aKaeyol & "'"
            sStr = sStr & "                         )"
            
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"         '<< 인문
                    sStr = sStr & "            AND SEL1 > ' '"
                Case "02"               '<< 자연
                    sStr = sStr & "            AND SEL3 > ' '"
            End Select
            
            sStr = sStr & "             )"
            sStr = sStr & "       ) B"
            sStr = sStr & "  WHERE A.ACID  = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND A.ACID  = B.ACID(+)"
            sStr = sStr & "    AND A.LSNCD = B.SEL_CLASS(+)"
            sStr = sStr & "  GROUP BY A.ACID, A.LSNCD"
    sStr = sStr & "        )"
    sStr = sStr & "  ORDER BY ACID, LSNCDNM "
    
    
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
'    ' KAEYOL
'        sTmp = aKaeyol
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' LSNTYPE
'        sTmp = aLsnType
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' KAEYOL
'        sTmp = aKaeyol
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' LSNTYPE
'        sTmp = aLsnType
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'    ' ACID
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
                sprClass.MaxRows = sprClass.MaxRows + 1
                sprClass.Row = sprClass.MaxRows:    sprClass.RowHeight(sprClass.Row) = nRowHeight
                
                sprClass.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprClass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprClass.Col = sprClass.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprClass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("LSNCAPA")) = True Then nTmp = CLng(.Fields("LSNCAPA"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_OK")) = True Then nTmp = CLng(.Fields("SEL_OK"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp):   sprClass.ForeColor = basModule.SectionColor1
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("PROC_NO")) = True Then nTmp = CLng(.Fields("PROC_NO"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp):   sprClass.ForeColor = basModule.SectionColor2
                
                sprClass.SetCellBorder sprClass.Col, sprClass.Row, sprClass.Col, sprClass.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL1")) = True Then nTmp = CLng(.Fields("SEL1"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL2")) = True Then nTmp = CLng(.Fields("SEL2"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL3")) = True Then nTmp = CLng(.Fields("SEL3"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL4")) = True Then nTmp = CLng(.Fields("SEL4"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL5")) = True Then nTmp = CLng(.Fields("SEL5"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL6")) = True Then nTmp = CLng(.Fields("SEL6"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL7")) = True Then nTmp = CLng(.Fields("SEL7"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL8")) = True Then nTmp = CLng(.Fields("SEL8"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                
                Select Case Trim(Right(cboKaeyol.Text, 30))
                    Case "01"       '<< 인문
                        sprClass.Col = sprClass.Col + 1
                            nTmp = 0: If IsNumeric(.Fields("SEL9")) = True Then nTmp = CLng(.Fields("SEL9"))
                                Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                        sprClass.Col = sprClass.Col + 1
                            nTmp = 0: If IsNumeric(.Fields("SEL10")) = True Then nTmp = CLng(.Fields("SEL10"))
                                Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                        sprClass.Col = sprClass.Col + 1
                            nTmp = 0: If IsNumeric(.Fields("SEL11")) = True Then nTmp = CLng(.Fields("SEL11"))
                                Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                End Select
                
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_X2")) = True Then nTmp = CLng(.Fields("SEL_X2"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_N1")) = True Then nTmp = CLng(.Fields("SEL_N1"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_N2")) = True Then nTmp = CLng(.Fields("SEL_N2"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_N3")) = True Then nTmp = CLng(.Fields("SEL_N3"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                sprClass.Col = sprClass.Col + 1
                    nTmp = 0: If IsNumeric(.Fields("SEL_N4")) = True Then nTmp = CLng(.Fields("SEL_N4"))
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 99999, ",", nTmp)
                
    
'                sprClass.Col = sprClass.MaxCols
'                    Call basFunction.Set_SprType_ChkBox(sprClass)
'                    sprClass.Value = 0
                    
                sprClass.SetCellBorder sprClass.Col, sprClass.Row, sprClass.Col, sprClass.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                .MoveNext
            Next nRec
            
            sprClass.Row = 1:       sprClass.Row2 = sprClass.MaxRows
            sprClass.Col = 1:       sprClass.Col2 = sprClass.MaxCols
            sprClass.BlockMode = True
                sprClass.BackColor = basModule.WhiteColor
                sprClass.BackColorStyle = BackColorStyleUnderGrid
            sprClass.BlockMode = False

            sprClass.ColsFrozen = 5
            
        '>> spread lock
            sprClass.Row = 1:       sprClass.Row2 = sprClass.MaxRows
            sprClass.Col = 1:       sprClass.Col2 = 4
            sprClass.BlockMode = True
                sprClass.Lock = True
                sprClass.Protect = True
            sprClass.BlockMode = False
            
            sprClass.Row = 1:       sprClass.Row2 = sprClass.MaxRows
            sprClass.Col = 6:       sprClass.Col2 = sprClass.MaxCols
            sprClass.BlockMode = True
                sprClass.Lock = True
                sprClass.Protect = True
            sprClass.BlockMode = False
            
        End If
    End With
    
    MsgBox "반 조회하였습니다.", vbInformation + vbOKOnly, "학생조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "반 등록할 학생 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"

End Sub















'##########################################################################################################
'## 반설정 알고리즘
'##########################################################################################################

Private Sub cmdProcClass_Click()
    
    Dim nRow            As Long
    Dim nCol            As Long
    
    Dim sClass          As String           ' 반명
    Dim nLimit          As Long             ' 인원수
    Dim nSubj           As Long             ' 과목수
    
    Dim nTGwamokCnt     As Long
    
    Dim nTotinwon       As Long
    Dim sHeader         As String
    Dim sTmp            As String
    Dim nTmp            As Long
    Dim nMaxStdinwon    As Long
    
    If sprClass.MaxRows = 0 Then
        MsgBox "반을 조회하세요.", vbExclamation + vbOKOnly, "반 설정"
        Exit Sub
    End If
    
    
    
'<<  선택학생 -> sprClassDet에 과목별 count   >>
    Select Case Trim(Right(txtKaeyol.Text, 30))
        Case "01", "03"         '<< 인문계 : 11 과목
            With sprClassDet
                .MaxCols = 0
                .MaxRows = 0
                
                .MaxCols = 18
                .MaxRows = 2
                
                .Row = SpreadHeader
                
                    sTmp = "국사":      .Col = 8:           .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "윤리":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "경제":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "한근":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "세계사":    .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "경지":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "한지":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "정치":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "사문":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "법사":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "세지":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                                
            End With
            
            '## 과목 선택인원
            For nCol = 8 To (11 + 8 - 1) Step 1         '<< 11 과목
                nTotinwon = 0
            
                sprSTD.Col = nCol
                sprSTD.Row = SpreadHeader
                    sHeader = Trim(sprSTD.Text)         '<< 헤더랑 동일한 내용이면 count + 1
                    
                For nRow = 1 To sprSTD.MaxRows Step 1
                    sprSTD.Row = nRow
                    sprSTD.Col = sprSTD.MaxCols
                    If sprSTD.value = 1 Then
                    
                        sprSTD.Col = nCol
                        sprSTD.Row = nRow
                            sTmp = Trim(sprSTD.Text)
                        
                        If StrComp(sHeader, sTmp, vbTextCompare) = 0 Then
                            nTotinwon = nTotinwon + 1
                            
                        End If
                        
                    End If
                Next nRow
                
                '## total 인원 체크
                sprClassDet.Row = 1
                sprClassDet.Col = nCol
                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -9999999, 9999999, ",", nTotinwon)
                
            Next nCol
    
        Case "02"               '<< 자연계 : 8 과목
            With sprClassDet
                .MaxCols = 0
                .MaxRows = 0
                
                .MaxCols = 15
                .MaxRows = 2
                
                .Row = SpreadHeader
                
                    sTmp = "물1":       .Col = 8:           .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "화1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "생1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "지1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "물2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "화2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "생2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "지2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
            End With
            
            '## 과목 선택인원
            For nCol = 8 To (8 + 8 - 1) Step 1          '<< 8과목
                nTotinwon = 0
            
                sprSTD.Col = nCol
                sprSTD.Row = SpreadHeader
                    sHeader = Trim(sprSTD.Text)         '<< 헤더랑 동일한 내용이면 count + 1
                    
                For nRow = 1 To sprSTD.MaxRows Step 1
                    sprSTD.Row = nRow
                    sprSTD.Col = sprSTD.MaxCols
                    If sprSTD.value = 1 Then
                    
                        sprSTD.Col = nCol
                        sprSTD.Row = nRow
                            sTmp = Trim(sprSTD.Text)
                        
                        If StrComp(sHeader, sTmp, vbTextCompare) = 0 Then
                            nTotinwon = nTotinwon + 1
                            
                        End If
                        
                    End If
                Next nRow
                
                '## total 인원 체크
                sprClassDet.Row = 1
                sprClassDet.Col = nCol
                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -9999999, 9999999, ",", nTotinwon)
                
            Next nCol
            
    End Select
    
    
    
    
    
'<< 매칭 >>
    With sprClass
        sClass = ""
        nLimit = 0
        nSubj = 0
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 2
                sClass = Trim(.Text)        '<< 반 명칭
            
            .Col = 5
                nLimit = .value
            
        '>> 정원 수와 가장 비슷하게 많은 학생이 신청한 과목의 학생 수를 구한다.
            If Select_Student(sClass, nLimit, nSubj) = True Then

            '>> 위에서 선택된 과목의 학생들에 대한 신청과목 수를 조사.
                Call Select_Order_Gwamok(nSubj)
                
            '>> 위에서 선택된 과목의 학생들중 두번째로 신청이 높은 과목의 학생들을 구한다.
                nTmp = Select_Sec_Order_Gwamok(nLimit, nSubj)
                
            
                If nTmp = 0 Then
                    '>> 두번째 신청이 높은 학생수가 정원보다 작을 경우
                    Call Make_Class_Less_OrdBok(sClass, nLimit)
                    
                Else
                
                    Call Make_Class_Great_OrdBok(sClass, nLimit, nTmp)
                    
                End If
                
                Call ReAction_sprSTD        '<< 설정되지 않은 학생에 대한 반 초기화
                
            End If
        Next nRow
    End With
        
        
    With sprClass
        ' 반의 학생수를 적용함.
        For nRow = 1 To .MaxRows Step 1
            nMaxStdinwon = 0
            
            .Row = nRow
            .Col = 2
                sClass = .Text                      '<< 반정보를 적용
            
            For nCol = 6 To .MaxCols Step 1         '<< column 정보 변경 : sprClass는 총원/ 선택/ 남은인원에 대한 정보가 있으므로 column 시작이 6부터
                
                sTmp = Set_Minus_Class_inwon(sClass, nCol + 2)      '<< 해당 과목에 대한 학생의 수를 샌다. 선택학생 수. (sprClassDet는 시작이 8부터, sprClass는 시작이 6부터)
                    
                If IsNumeric(sTmp) = True Then
                    nTmp = CLng(sTmp)
                    
                    If nMaxStdinwon <= nTmp Then nMaxStdinwon = nTmp
                    
                    .Col = nCol
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 999999, ",", nTmp)
                    
                End If
                
            Next nCol
            
            '## 최대인원을 삭제 : 선택인원/ 남은인원 계산
            .Row = nRow
            .Col = 5
            If IsNumeric(.Text) = True Then
                If .value > 0 Then
                    .Col = 4:   nTmp = nMaxStdinwon
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 999999, ",", nTmp)
                    .Col = 5
                        nTmp = .value - nMaxStdinwon
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 999999, ",", nTmp)
                End If
            End If
            
        Next nRow
        
    End With
    
End Sub


'<< 해당 과목에 대한 학생의 수를 샌다. 선택학생 수.
Private Function Set_Minus_Class_inwon(ByVal aClass As String, ByVal aCol As Long) As Long
    Dim nRow        As Long
    Dim nCnt        As Long
    
    nCnt = 0
    
    With sprSTD
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2
            
            If StrComp(Trim(.Text), aClass, vbTextCompare) = 0 Then             ' 이 반의 학생이 맞다면
                .Col = aCol                 ' 선택과목을 확인
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< 인문계 : 11 과목
                        
                        If StrComp(Trim(.Text), "국사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "윤리", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "경제", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "한근", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "세계사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "경지", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "한지", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "정치", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "사문", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "법사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "세지", vbTextCompare) = 0 Then
                           
                            nCnt = nCnt + 1     '<< 선택했다면 선택학생수 한명 증가.
                            
                        End If
                        
                    Case "02"
                            
                        If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then
                           
                            nCnt = nCnt + 1     '<< 선택했다면 선택학생수 한명 증가.
                            
                        End If
                End Select
                
            End If
        Next nRow
    End With
    
    Set_Minus_Class_inwon = nCnt

End Function


'<< 설정되지 않은 학생에 대한 반 초기화
Private Sub ReAction_sprSTD()
    Dim nRow        As Long
    Dim nCol        As Long
    
    ' 반을 다 설정하고 난 후 나머지 학생들에 대한 처리.
    ' 선택한 학생(maxcols가 0으로 Setting된 학생들)을 다시 널로 초기화.
    ' maxcols가 0이라 함은 선택은 됐으나 반에 속해지지 못한 학생들을 의미함.
    With sprSTD
        .Col = .MaxCols
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2
                If .Text = "0" Then .Text = ""
            
        Next nRow
    End With

End Sub

'>> 두번째 선택한 과목의 학생수가 정원수보다 많을때
'   두번째 선택한 과목을 선택하여 그 학생들을 위에서부터 정원수까지 배정.
Private Sub Make_Class_Great_OrdBok(ByVal aClass As String, ByVal nLimit As Long, ByVal nSubj As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nC          As Long
    Dim nTmp        As Long
    
    With sprSTD
        nC = 0
        
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = .MaxCols - 2
            
            If StrComp(Trim(.Text), "0", vbTextCompare) = 0 Then            '<< 첫번째 선택과목이 지금 설정된 과목이라면... 선택된 학생들이라면..
                
                .Col = nSubj                '두번째 선택과목을 선택
                
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< 인문계 : 11 과목
                        
                        If StrComp(Trim(.Text), "국사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "윤리", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "경제", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "한근", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "세계사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "경지", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "한지", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "정치", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "사문", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "법사", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "세지", vbTextCompare) = 0 Then
                        
                            nC = nC + 1
                            
                            If nC > nLimit Then Exit Sub
                            
                            .Col = .MaxCols - 2
                                .Text = aClass          '<< "0" 을 반명으로 대치
                
                            .Col = .MaxCols
                                .value = 0          ' 선택해제
                            
                            .Row2 = .Row
                            .Col = 1:   .Col2 = .MaxCols
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            
                            For nCol = 8 To (11 + 8 - 1) Step 1
                                .Col = nCol
                                
                                If StrComp(Trim(.Text), "국사", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "윤리", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "경제", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "한근", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "세계사", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "경지", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "한지", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "정치", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "사문", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "법사", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "세지", vbTextCompare) = 0 Then
                                         
                                    With sprClassDet
                                        .Row = 1
                                        .Col = nCol
                                        
                                            If IsNumeric(.Text) = False Then
                                                nTmp = 0
                                            Else
                                                nTmp = .value - 1
                                            End If
                                            Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)
                                            
                                    End With
                                End If
                            Next nCol
                            
                        End If
                        
                        
                    Case "02"
                            
                        If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then
                           
                            nC = nC + 1
                            
                            If nC > nLimit Then Exit Sub
                            
                            .Col = .MaxCols - 2
                                .Text = aClass          '<< "0" 을 반명으로 대치
                
                            .Col = .MaxCols
                                .value = 0          ' 선택해제
                            
                            .Row2 = .Row
                            .Col = 1:   .Col2 = .MaxCols
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            
                            For nCol = 8 To (8 + 8 - 1) Step 1
                                .Col = nCol
                                
                                If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then
                                         
                                    With sprClassDet
                                        .Row = 1
                                        .Col = nCol
                                        
                                            If IsNumeric(.Text) = False Then
                                                nTmp = 0
                                            Else
                                                nTmp = .value - 1
                                            End If
                                            Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)
                                            
                                    End With
                                End If
                            Next nCol
                           
                        End If
                    
                        
                End Select
                
            End If
        Next nRow
    End With
End Sub


'>> 두번째 선택과목수가 해당 반의 정원보다 적을 경우, 첫번째 선택과목 선택학생들중
'   위에서부터 정원수대로 잘라 반에 배정한다.
Private Sub Make_Class_Less_OrdBok(ByVal aClass As String, ByVal nLimit As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nC          As Long
    Dim nTmp        As Long
    
    With sprSTD
        nC = 0
        
        For nRow = 1 To .MaxRows
            .Row = nRow
            .Col = .MaxCols - 2
            
            If StrComp(.Text, "0", vbTextCompare) = 0 Then
            
                nC = nC + 1
                If nC > nLimit Then Exit Sub
                
                .Text = aClass          '<< "0" 을 반명으로 대치
                
                .Col = .MaxCols
                    .value = 0          ' 선택해제
                
                .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< 인문계 : 11 과목
                        
                        For nCol = 8 To (11 + 8 - 1) Step 1         '< 2007.12.17
                        
                            .Col = nCol
                            
                            If StrComp(Trim(.Text), "국사", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "윤리", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "경제", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "한근", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "세계사", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "경지", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "한지", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "정치", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "사문", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "법사", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "세지", vbTextCompare) = 0 Then
                            
                                With sprClassDet
                                    .Row = 1
                                    .Col = nCol
                                    
                                        If IsNumeric(.Text) = False Then
                                            nTmp = 0
                                        Else
                                            nTmp = .value - 1
                                        End If
                                        Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)
                                        
                                End With
                            
                            End If
                                
                        Next nCol
                    Case "02"               '<< 자연계 : 8 과목
                        
                        For nCol = 8 To (8 + 8 - 1) Step 1          '< 2007.12.17
                        
                            .Col = nCol
                            
                            If StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "지2", vbTextCompare) = 0 Then
                                                   
                                With sprClassDet
                                    .Row = 1
                                    .Col = nCol
                                    
                                        If IsNumeric(.Text) = False Then
                                            nTmp = 0
                                        Else
                                            nTmp = .value - 1
                                        End If
                                        Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)
                                        
                                End With
                                
                            End If
                                
                        Next nCol
                        
                        
                End Select
                        
            End If
        Next nRow
    End With
End Sub










'>> 학생의 두번째 선택이 많은 과목을 선택한다.
'   nLimit 는 지금 선택된 반의 정원 수, 1는 그 선택과목에 대한 컬럼 수
Private Function Select_Sec_Order_Gwamok(ByVal nLimit As Long, ByVal nSubj As Long) As Long
    Dim nCol    As Integer
    Dim iSubj   As Integer
    Dim nTmp    As Long
    
    iSubj = 0
    
    With sprClassDet
        .Row = .MaxRows
        
        For nCol = 8 To .MaxCols
            .Col = nCol
                If IsNumeric(.Text) = False Then
                    nTmp = 0
                Else
                    nTmp = CLng(.Text)
                End If
                
            If .Col <> nSubj Then               ' 학생의 첫번째 선택과목 제외
                If nLimit < nTmp Then           ' 반의 정원 수 보다 많다면 OK(그거 선택)
                    iSubj = nCol
                    Select_Sec_Order_Gwamok = iSubj
                    
                    Exit Function
                    
                End If
            End If
        Next nCol
    End With
    
    Select_Sec_Order_Gwamok = iSubj
End Function


'>> 위에서 선택된 과목의 학생들에 대한 신청과목 수를 조사.
Private Sub Select_Order_Gwamok(ByVal nSubj As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sHeader     As String
    Dim sTmp        As String
    Dim nTmp        As Long
    
    With sprClassDet                '<< 과목선택했던 내용을 모두 초기화
        .Row = .MaxRows             ' 2번째 열을 초기화
        
        For nCol = 1 To .MaxCols
            .Col = nCol:        .Text = ""
        Next nCol
    End With
    
    ' 맨 위 스프레드 처음 학생부터 시작.
    ' maxcols는 반명을 표시하게 되어 있는데 그게 널이면 아직 반이 설정되지 않았다고 보고.
    ' 해당 과목을 선택한 학생인지 확인한다. nsubj는 컬럼수으로 적용했는데 나중에 바꿔도 됨.
    With sprSTD
        For nRow = 1 To .MaxRows Step 1
            .Row = SpreadHeader
            .Col = nSubj
                sHeader = Trim(.Text)
            
            .Row = nRow
            .Col = .MaxCols
            
            If .value = 1 Then
            
                .Row = nRow
                .Col = nSubj
                    sTmp = Trim(.Text)
                
                If StrComp(sHeader, sTmp, vbTextCompare) = 0 Then       ' 선택한 과목이라면
                
                    .Col = .MaxCols - 2
                    .Text = 0                                           ' 일단 반은 초기값 0으로, 0은 선택된 학생이라는 표시
                    
                    With sprClassDet
                        .Row = .MaxRows
                        .Col = nSubj
                        
                            If IsNumeric(.Text) = False Then
                                nTmp = 1
                            Else
                                nTmp = .value + 1
                            End If
                            Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)
                    End With
                    
                    
                    ' 그 외의 과목들의 수를 조사하여 해당 학생에 대한 선택과목을 계산한다.
                    ' 즉 그 학생의 나머지 과목들에 대해서도 적용하여 sprClassDet의 정보를 바꿔 정확한 정보로 만드는거지.
                    Select Case Trim(Right(txtKaeyol.Text, 30))
                        Case "01", "03"         '<< 인문계 : 11 과목
                            For nCol = 8 To (11 + 8 - 1) Step 1
                                .Col = nCol
                                
                                '지금 보고있는 과목(아까 정원수로 해서 찾은 그 과목)이 아닌 다른 과목일경우,
                                '그리고 그 다른 과목을 이 학생이 선택한 경우라면.. sprClassDet에 정보 업데이트해야지. 제대로
                            
                                If .Col <> nSubj And _
                                    (StrComp(Trim(.Text), "국사", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "윤리", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "경제", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "한근", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "세계사", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "경지", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "한지", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "정치", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "사문", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "법사", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "세지", vbTextCompare) = 0) Then
                                                   
                                        With sprClassDet
                                            .Row = .MaxRows
                                            .Col = nCol
                                            
                                                If IsNumeric(.Text) = False Then
                                                    nTmp = 1
                                                Else
                                                    nTmp = .value + 1
                                                End If
                                                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)    '선택과목 학생 수 증가
                                        End With

                                End If
                                
                            Next nCol
                        
                        Case "02"               '<< 자연계 : 8 과목
                        
                            For nCol = 8 To (8 + 8 - 1) Step 1
                                .Col = nCol
                                
                                '지금 보고있는 과목(아까 정원수로 해서 찾은 그 과목)이 아닌 다른 과목일경우,
                                '그리고 그 다른 과목을 이 학생이 선택한 경우라면.. sprClassDet에 정보 업데이트해야지. 제대로
                            
                                If .Col <> nSubj And _
                                    (StrComp(Trim(.Text), "물1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "화1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "생1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "지1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "물2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "화2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "생2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "지2", vbTextCompare) = 0) Then
                                                   
                                        With sprClassDet
                                            .Row = .MaxRows
                                            .Col = nCol
                                            
                                                If IsNumeric(.Text) = False Then
                                                    nTmp = 1
                                                Else
                                                    nTmp = .value + 1
                                                End If
                                                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)    '선택과목 학생 수 증가
                                        End With
                                    
                                End If
                                
                            Next nCol
                                
                    End Select
        
                End If
                
            End If
        Next nRow
        
    End With
End Sub


'>> 정원 수와 가장 비슷하게 많은 학생이 신청한 과목의 학생 수를 구한다.
Private Function Select_Student(ByVal sBan As String, ByVal nLimit As Integer, ByRef nSubj As Long) As Boolean
    Dim nCols       As Long
    Dim nTmp        As Long
    Dim nC          As Long
    
    Dim bChk        As Boolean

    bChk = False

    nC = 0
    ' sprClass 는 총 학생수가 나와있는 스프레드.
    ' 그 스프레드에서 현재 가지고 온 nLimit라는 변수= 반의 정원.
    ' 반의 정원보다 크고, 그 큰 놈들중 가장 작은 과목을 선택하여 그 해당 과목신청자를 찾음
    With sprClassDet
        .Row = 1

        For nCols = 8 To .MaxCols Step 1        ' 각 과목별 전체 내역을 모두 읽음. : 과목은 8번째 줄부터 시작
            .Col = nCols
            
            If IsNumeric(.Text) = True Then
                nTmp = .value
            Else
                nTmp = 0
            End If
            
            If nTmp > nLimit Then               ' 만약에 정원수보다 해당과목을 선택한 학생수가 더 많다면 오케이
                bChk = True
                
                If nC = 0 Then
                    nC = val(.Text)             ' nC =  최소의 수를 구하기 위한 저장공간.
                                                ' 이 변수를 사용하여 변수를 계속 비교해가면서 정원보다는 많고, 그중 가장 작은 수를 선택하게 됨.
                    nSubj = nCols
                    
                ElseIf nC > val(.Text) Then
                    nC = val(.Text)
                    nSubj = nCols
                    
                End If
            End If
        Next nCols
    End With

    If bChk = False Then                ' 만약 반의 정원 수보다 작은 인원의 선택과목밖에 없을경우 처리
                                        ' 정원수보다 다 작다면.. 어떻게 하면 좋을지..
        ' 여기는 아직 처리하지 않음.
    End If

    Select_Student = bChk
End Function
























'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%% 반 등록하기
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'>> 반등록시 UPDATE 만 있습니다.


Private Sub cmdinput_Class_Click()

    Dim nRow        As Long
    Dim nChk        As Long
    Dim uClass()    As tClass

    Dim nRec        As Long
    Dim sClassNM    As String
    Dim sTmp        As String

    Dim ninClass()  As Long         ' 저장할 행
    Dim nC          As Long

    nChk = 0

    With sprClass
        If .MaxRows = 0 Then
            MsgBox "반 설정을 조회하세요.", vbExclamation + vbOKOnly, "반 등록하기"
            Exit Sub
        End If

        ReDim uClass(.MaxRows) As tClass

        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 1:           uClass(nRow).CLSCD = Trim(.Text)
            .Col = .Col + 1:    uClass(nRow).CLSNM = Trim(.Text)
        Next nRow
    End With

    With sprSTD
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2                 '< 반명
            If Trim(.Text) > " " Then
                nChk = nChk + 1
                Exit For
            End If
        Next nRow

        If nChk = 0 Then
            MsgBox "처리된 반이 없습니다.", vbExclamation + vbOKOnly, "반 등록하기"
            Exit Sub
        End If

        ReDim ninClass(0) As Long
        nC = 0

        For nRow = 1 To .MaxRows Step 1
            nChk = 0                        '<< 등록가능 체크

            .Row = nRow
            .Col = .MaxCols - 2             '< 반명
                sClassNM = Trim(.Text)

            If sClassNM > " " Then
                For nRec = 1 To UBound(uClass) Step 1
                    If StrComp(sClassNM, uClass(nRec).CLSNM, vbTextCompare) = 0 Then

                        sTmp = uClass(nRec).CLSCD
                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", LenB(sTmp), sTmp)

                        nChk = nChk + 1

                        '## 저장할 행
                        nC = nC + 1
                        ReDim Preserve ninClass(nC) As Long
                        ninClass(nC) = .Row

                    End If
                Next nRec

                If nChk = 0 Then
                    MsgBox Trim(CStr(.Row)) & "행" & vbCrLf & "반 명이 잘못되었으니 확인하십시요.", vbExclamation + vbOKOnly, "반 등록하기"
                    Exit Sub
                End If
            End If

        Next nRow
    End With

    If UBound(ninClass) > 0 Then
        If input_Class_Data(ninClass) = True Then
            MsgBox "반 등록하였습니다.", vbInformation + vbOKOnly, "반 등록하기"
        Else
            MsgBox "반 등록을 못하였습니다.", vbCritical + vbOKOnly, "반 등록하기"
        End If
    Else
        MsgBox "처리할 내용이 없습니다.", vbExclamation + vbOKOnly, "반 등록하기"
    End If

End Sub


'## 등록하기
Private Function input_Class_Data(ByRef ainClass() As Long) As Boolean
    Dim bRet        As Boolean

    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim ni          As Long

    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    Dim nTotExe     As Long

    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long

    Dim nRow        As Long
    Dim sClassCD    As String
    Dim sSchNO      As String

    Dim bUpChks     As Boolean

    bRet = False

    On Error GoTo ErrStmt

    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    nTotExe = 0

    For nRec = 1 To UBound(ainClass) Step 1

            sprSTD.Row = ainClass(nRec)
            sprSTD.Col = sprSTD.MaxCols - 2             '< 반명
                sClassCD = Trim(sprSTD.Text)

            sprSTD.Row = ainClass(nRec)
            sprSTD.Col = 1                              '< 학생코드 (시스템)
                sSchNO = Trim(sprSTD.Text)

        sStr = ""
        sStr = sStr & " UPDATE CLTTL01TB"
        sStr = sStr & "    SET SEL_CLASS = '" & sClassCD & "'"
        sStr = sStr & "  WHERE SCHNO = '" & sSchNO & "'"
        sStr = sStr & "    AND ACID  = '" & Trim(basModule.SchCD) & "'"

  

'    '>> 반코드
'        sprSTD.Row = ainClass(nRec)
'        sprSTD.Col = sprSTD.MaxCols - 2
'            sClassCD = Trim(sprSTD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SEL_CLASS", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'
'    '>> 학생
'        sprSTD.Row = ainClass(nRec)
'        sprSTD.Col = 1
'            sTmp = Trim(sprSTD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'
'    '>> 학원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam


        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30

        nExe = 0
        DBCmd.Execute nExe, , -1

        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop

        If nExe = 1 Then

        '<< 아래의 부분은 PROCEDURE 사용

 

            '>> 학원코드
            sTmp = Trim(basModule.SchCD)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> 데이터 등록
            DBCmd.CommandType = adCmdStoredProc
            DBCmd.CommandText = "PROC_CLASS.CHG_SEL_CLASS_DATA"
            DBCmd.CommandTimeout = 30

            DBCmd.Execute

            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop

            nTotExe = nTotExe + 1

        End If

    Next nRec

    If nTotExe = UBound(ainClass) Then
        basDataBase.DBConn.CommitTrans
        input_Class_Data = True                 '<< OK
    Else
        basDataBase.DBConn.RollbackTrans
        input_Class_Data = False                '<< FAIL
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing

    Exit Function

ErrStmt:
    basDataBase.DBConn.RollbackTrans

    Set DBCmd = Nothing
    Set DBParam = Nothing

    input_Class_Data = False

End Function










'**********************************************************************************************************
'** 삭제하기
'**********************************************************************************************************


Private Sub sprClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            With sprClass
                .Row = .ActiveRow
                .DeleteRows .Row, 1
                .MaxRows = .MaxRows - 1
            End With
    End Select
End Sub

Private Sub sprClass_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sTmp    As String
    
    If Row < 1 Then Exit Sub
    
    With sprClass
        If .MaxRows < 1 Then Exit Sub
        If .Tag = "" Then .Tag = "1"
    
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
        .BackColor = basModule.SelectColor2
        .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub


'<< 선택반 내역 삭제하기
Private Sub cmdDeleteClass_Click()
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim ni          As Long

    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    Dim nTotExe     As Long
    Dim nChk1       As Long

    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long

    Dim nRow        As Long
    Dim sClassCD    As String

    Dim bUpChks     As Boolean

    On Error GoTo ErrStmt

    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    nTotExe = 0
    nChk1 = 0

    For nRec = 1 To sprClass.MaxRows Step 1

        sprClass.Row = nRec
        sprClass.Col = 1

        If sprClass.BackColor = basModule.SelectColor2 Then
            nChk1 = nChk1 + 1

            sprClass.Row = nRec
                sprClass.Col = 1                '< 학생코드
                    sClassCD = Trim(sprClass.Text)

            sStr = ""
            sStr = sStr & " UPDATE CLTTL01TB"
            sStr = sStr & "    SET SEL_CLASS = ''   "       '<< class 없앰
            sStr = sStr & "  WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND SEL_CLASS = '" & sClassCD & "'"



    '    '>> 학원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> 반
    '        sTmp = sClassCD
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("SEL_CLASS", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30

            nExe = 0
            DBCmd.Execute nExe, , -1

            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop

            If nExe = 0 Then
                MsgBox "삭제될 정보가 없습니다.", vbExclamation + vbOKOnly, "선택반 등록내역 삭제하기"
                basDataBase.DBConn.RollbackTrans
                
                Set DBCmd = Nothing
                Set DBParam = Nothing
                
                Call cmdClass_Click         '<< 반 조회
                Call cmdFindStd_Click       '<< 학생조회
    
            ElseIf nExe > 0 Then

            '<< 아래의 부분은 PROCEDURE 사용




                '>> 학원코드
                sTmp = Trim(basModule.SchCD)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

                '>> 데이터 등록
                DBCmd.CommandType = adCmdStoredProc
                DBCmd.CommandText = "PROC_CLASS.CHG_SEL_CLASS_DATA"
                DBCmd.CommandTimeout = 30

                DBCmd.Execute

                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop

                nTotExe = nTotExe + 1

            End If


        End If
    Next nRec

    If nTotExe = nChk1 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "선택반 등록내역 삭제하기"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "삭제시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "선택반 등록내역 삭제하기"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Call cmdClass_Click         '<< 반 조회
    Call cmdFindStd_Click       '<< 학생조회

    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "선택반 등록내역 삭제하기"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub

'<< 선택학생 반 삭제하기
Private Sub cmdDelStdClass_Click()
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim ni          As Long

    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    Dim nTotExe     As Long
    Dim nChk1       As Long

    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long

    Dim nRow        As Long
    Dim sSchNO      As String

    Dim bUpChks     As Boolean

    On Error GoTo ErrStmt

    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    nTotExe = 0
    nChk1 = 0
    
    For nRec = 1 To sprSTD.MaxRows Step 1

        sprSTD.Row = nRec
        sprSTD.Col = sprSTD.MaxCols

        If sprSTD.value = 1 Then
            nChk1 = nChk1 + 1

            sprSTD.Row = nRec
                sprSTD.Col = 1              '< 학생코드
                    sSchNO = Trim(sprSTD.Text)

            sStr = ""
            sStr = sStr & " UPDATE CLTTL01TB"
            sStr = sStr & "    SET SEL_CLASS = ''   "       '<< class 없앰
            sStr = sStr & "  WHERE SCHNO = '" & sSchNO & "'"
            sStr = sStr & "    AND ACID  = '" & Trim(basModule.SchCD) & "'"



    '    '>> 학생
    '        sprSTD.Row = nRec
    '        sprSTD.Col = 1
    '            sTmp = Trim(sprSTD.Text)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> 학원
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30

            nExe = 0
            DBCmd.Execute nExe, , -1

            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop

            If nExe = 0 Then
                MsgBox "선택학생의 반 내역이 없습니다.", vbExclamation + vbOKOnly, "선택반 등록내역 삭제하기"
                basDataBase.DBConn.RollbackTrans
            
                Set DBCmd = Nothing
                Set DBParam = Nothing
                
                Call cmdClass_Click         '<< 반 조회
                Call cmdFindStd_Click       '<< 학생조회
                
            ElseIf nExe > 0 Then

            '<< 아래의 부분은 PROCEDURE 사용




                '>> 학원코드
                sTmp = Trim(basModule.SchCD)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

                '>> 데이터 등록
                DBCmd.CommandType = adCmdStoredProc
                DBCmd.CommandText = "PROC_CLASS.CHG_SEL_CLASS_DATA"
                DBCmd.CommandTimeout = 30

                DBCmd.Execute

                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop

                nTotExe = nTotExe + 1

            End If
            
        End If
    Next nRec

    If nTotExe = nChk1 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "선택반 등록내역 삭제하기"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "삭제시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "선택반 등록내역 삭제하기"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Call cmdClass_Click         '<< 반 조회
    Call cmdFindStd_Click       '<< 학생조회
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "선택반 등록내역 삭제하기"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub








'>> 전체 정렬 : 2007.12.17
'       ## SORT ORDER ##
Private Sub cmdSort_Click()
    Call spread_Sort("CMD")
    
End Sub

Private Sub spread_Sort(Optional aClick As String)
    Dim ni      As Integer
    Dim nj      As Integer
    Dim nC      As Integer
    
    Dim sSort   As String
    Dim sR      As String
    
    Dim sDivs() As String
    Dim sDivC() As String
    
    nC = 0
    sSort = ""
    
    With sprSTD
        For ni = 1 To 6 Step 1
            For nj = 1 To 6 Step 1
                If fpSort(nj).value = ni Then
                    nC = nC + 1
                    
                    Select Case nj
                        Case 1                      '<< 언어
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 2                      '<< 수리
                            .SortKey(nC) = 5
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 3                      '<< 외국어
                            .SortKey(nC) = 6
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 4                      '<< 합계
                            .SortKey(nC) = 7
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 5                      '<< MU_TYPE
                            .SortKey(nC) = .MaxCols - 1
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 6                      '<< 수험번호
                            .SortKey(nC) = 3
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                    End Select
                    
                End If
            Next nj
        Next ni
        
        For ni = 0 To 2 Step 1
            If Trim(Right(cboGwamok(ni).Text, 30)) <> "X" Then
                nC = nC + 1
                
                .SortKey(nC) = CLng(Trim(Right(cboGwamok(ni).Text, 30)))
                .SortKeyOrder(nC) = SortKeyOrderAscending
            End If
        Next ni
        
        If aClick <> "CMD" Then         '< 버튼 클릭이 아닌경우
            Select Case .ActiveCol
                Case 4 To 7
                
                Case Else
                    .SortKey(nC + 1) = .ActiveCol
                    .SortKeyOrder(nC + 1) = SortKeyOrderDescending
            End Select
        End If
        
        If nC = 0 Then
            .SortKey(1) = .ActiveCol
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If
        
        .Sort -1, -1, -1, -1, SortByRow
        
    End With
    
End Sub











