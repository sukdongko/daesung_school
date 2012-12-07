VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form STD020 
   Caption         =   "입학사정 >> 합격생 등록"
   ClientHeight    =   9975
   ClientLeft      =   3810
   ClientTop       =   3000
   ClientWidth     =   15855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   15855
   Begin VB.Frame Frame4 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame4"
      Height          =   8325
      Left            =   30
      TabIndex        =   42
      Top             =   1140
      Width           =   15045
      Begin VB.Frame Frame5 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame5"
         Height          =   8265
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   14985
         Begin EditLib.fpLongInteger fpSort 
            Height          =   315
            Index           =   9
            Left            =   3240
            TabIndex        =   68
            Top             =   240
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
            MaxValue        =   "9"
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
         Begin VB.TextBox Text1 
            Height          =   3135
            Left            =   2520
            TabIndex        =   66
            Text            =   "Text1"
            Top             =   2160
            Visible         =   0   'False
            Width           =   6255
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "합격취소"
            Height          =   500
            Left            =   480
            TabIndex        =   62
            Top             =   7680
            Width           =   1905
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "등록"
            Height          =   225
            Left            =   14010
            TabIndex        =   30
            Top             =   600
            Width           =   675
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "합격생 등록하기 (&S)"
            Height          =   500
            Left            =   12540
            TabIndex        =   32
            Top             =   7680
            Width           =   2000
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame6"
            Height          =   525
            Left            =   30
            TabIndex        =   50
            Top             =   30
            Width           =   9675
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
               Height          =   375
               Left            =   180
               TabIndex        =   16
               Top             =   90
               Width           =   645
            End
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   17
               Top             =   210
               Width           =   555
               _Version        =   196608
               _ExtentX        =   979
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
               MaxValue        =   "9"
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
               Left            =   2010
               TabIndex        =   18
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
               MaxValue        =   "9"
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
               Left            =   4050
               TabIndex        =   19
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
               MaxValue        =   "9"
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
               Left            =   7440
               TabIndex        =   23
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
               MaxValue        =   "9"
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
               Index           =   4
               Left            =   8250
               TabIndex        =   24
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
               MaxValue        =   "9"
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
               Left            =   9000
               TabIndex        =   25
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
               MaxValue        =   "9"
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
               Left            =   4860
               TabIndex        =   20
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
               MaxValue        =   "9"
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
               Index           =   7
               Left            =   5730
               TabIndex        =   21
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
               MaxValue        =   "9"
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
               Index           =   8
               Left            =   6570
               TabIndex        =   22
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
               MaxValue        =   "9"
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
            Begin VB.Label Label19 
               BackStyle       =   0  '투명
               Caption         =   "내신등급"
               Height          =   210
               Left            =   3150
               TabIndex        =   67
               Top             =   30
               Width           =   825
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "외국어"
               Height          =   210
               Index           =   3
               Left            =   6510
               TabIndex        =   61
               Top             =   15
               Width           =   615
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "수리"
               Height          =   210
               Index           =   2
               Left            =   5700
               TabIndex        =   60
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "언어"
               Height          =   210
               Index           =   1
               Left            =   4830
               TabIndex        =   59
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "선택"
               Height          =   210
               Index           =   0
               Left            =   8970
               TabIndex        =   44
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label11 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "계열"
               Height          =   210
               Left            =   8220
               TabIndex        =   45
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label10 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "합계"
               Height          =   210
               Left            =   7410
               TabIndex        =   46
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label9 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "유.무"
               Height          =   210
               Left            =   4050
               TabIndex        =   47
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label8 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "성명"
               Height          =   210
               Left            =   2070
               TabIndex        =   48
               Top             =   15
               Width           =   405
            End
            Begin VB.Label Label7 
               Alignment       =   1  '오른쪽 맞춤
               BackStyle       =   0  '투명
               Caption         =   "수험번호"
               Height          =   210
               Left            =   1050
               TabIndex        =   49
               Top             =   15
               Width           =   795
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
               Left            =   30
               TabIndex        =   51
               Top             =   165
               Width           =   645
            End
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   3
            Left            =   12960
            Style           =   2  '드롭다운 목록
            TabIndex        =   29
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   2
            Left            =   11880
            Style           =   2  '드롭다운 목록
            TabIndex        =   28
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   1
            Left            =   10800
            Style           =   2  '드롭다운 목록
            TabIndex        =   27
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   0
            Left            =   9720
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   240
            Width           =   1035
         End
         Begin FPSpread.vaSpread sprPass 
            Height          =   7035
            Left            =   30
            TabIndex        =   31
            Top             =   570
            Width           =   14925
            _Version        =   393216
            _ExtentX        =   26326
            _ExtentY        =   12409
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
            MaxCols         =   20
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD020.frx":0000
         End
         Begin VB.Label Label13 
            BackStyle       =   0  '투명
            Caption         =   ">> 합격학원 ------------------------------"
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
            Left            =   9750
            TabIndex        =   52
            Top             =   30
            Width           =   4515
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   30
      TabIndex        =   33
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   1005
         Left            =   30
         TabIndex        =   34
         Top             =   30
         Width           =   14985
         Begin VB.ComboBox cboSel2_Sch 
            Height          =   300
            Left            =   13530
            Style           =   2  '드롭다운 목록
            TabIndex        =   7
            Top             =   172
            Width           =   1395
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   4140
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   165
            Width           =   1005
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회하기 (&F)"
            Height          =   450
            Left            =   420
            TabIndex        =   0
            Top             =   30
            Width           =   1365
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00D2EAF5&
            Height          =   495
            Left            =   90
            TabIndex        =   41
            Top             =   450
            Width           =   2655
            Begin VB.OptionButton optPassN 
               BackColor       =   &H00D2EAF5&
               Caption         =   "합격처리할"
               Height          =   285
               Left            =   90
               TabIndex        =   8
               Top             =   150
               Width           =   1365
            End
            Begin VB.OptionButton optPassY 
               BackColor       =   &H00D2EAF5&
               Caption         =   "합격생만"
               Height          =   285
               Left            =   1500
               TabIndex        =   9
               Top             =   150
               Width           =   1035
            End
         End
         Begin VB.ComboBox cboHakwon 
            Height          =   300
            Left            =   2580
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   165
            Width           =   1005
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   9090
            TabIndex        =   5
            Text            =   "txtStdNM"
            Top             =   143
            Width           =   945
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3420
            Style           =   2  '드롭다운 목록
            TabIndex        =   10
            Top             =   600
            Width           =   1005
         End
         Begin EditLib.fpMask fpBirth_ymd 
            Height          =   345
            Left            =   11040
            TabIndex        =   6
            Top             =   150
            Width           =   1455
            _Version        =   196608
            _ExtentX        =   2566
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
         Begin EditLib.fpMask fpExmID_S 
            Height          =   345
            Left            =   6180
            TabIndex        =   3
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
            Left            =   7350
            TabIndex        =   4
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
         Begin VB.Frame Frame7 
            BackColor       =   &H00D2EAF5&
            Height          =   495
            Left            =   4620
            TabIndex        =   54
            Top             =   450
            Width           =   7755
            Begin EditLib.fpLongInteger fpTotS 
               Height          =   345
               Left            =   5640
               TabIndex        =   14
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
               Left            =   6750
               TabIndex        =   15
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
               Left            =   540
               TabIndex        =   11
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
               Left            =   2160
               TabIndex        =   12
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
               Left            =   3990
               TabIndex        =   13
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
            Begin VB.Label Label17 
               BackStyle       =   0  '투명
               Caption         =   "외국어            이상/"
               Height          =   210
               Left            =   3420
               TabIndex        =   58
               Top             =   180
               Width           =   1755
            End
            Begin VB.Label Label16 
               BackStyle       =   0  '투명
               Caption         =   "수리            이상/"
               Height          =   210
               Left            =   1770
               TabIndex        =   57
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label15 
               BackStyle       =   0  '투명
               Caption         =   "언어            이상/"
               Height          =   210
               Left            =   150
               TabIndex        =   56
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label6 
               BackStyle       =   0  '투명
               Caption         =   "합계             부터"
               Height          =   210
               Left            =   5220
               TabIndex        =   55
               Top             =   180
               Width           =   1995
            End
         End
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   14070
            TabIndex        =   63
            Top             =   600
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
         Begin VB.Label Label18 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "2지망 학원"
            Height          =   210
            Left            =   12570
            TabIndex        =   65
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '투명
            Caption         =   "조회인원"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   13350
            TabIndex        =   64
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "시험"
            Height          =   210
            Left            =   3690
            TabIndex        =   53
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "등록학원"
            Height          =   210
            Left            =   1560
            TabIndex        =   40
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "수험번호             부터             까지"
            Height          =   210
            Left            =   5460
            TabIndex        =   39
            Top             =   210
            Width           =   3075
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   10020
            TabIndex        =   38
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            Height          =   210
            Left            =   8100
            TabIndex        =   37
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계  열"
            Height          =   210
            Left            =   2370
            TabIndex        =   36
            Top             =   660
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
            Left            =   120
            TabIndex        =   35
            Top             =   150
            Width           =   2625
         End
      End
   End
End
Attribute VB_Name = "STD020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD020
'   모 듈  목 적 : 합격생 등록
'
'   작   성   일 : 2007/08/24
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private sini_Path       As String    '>> 대성학원
Private sChasuTimes     As String

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sSort       As String
    
    Dim sData       As String * 255
    Dim sGbn        As String
    Dim nRtn        As Long
    
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
        With sprPass
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
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
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            
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
                 Case "E"
                    .ListIndex = 9
            End Select
        End With
        
        With cboExmType
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
            .AddItem "무시험" & Space(30) & "0"
            .AddItem "유시험" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        '제2지망
        With cboSel2_Sch
            .Clear
            .AddItem "없음" & Space(30) & "X"
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            
            .ListIndex = 0
        End With
    
'>> 계열
        With cboKaeyol
            .Clear
            .AddItem "전체" & Space(30) & "ALL"
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
            If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then           '< 강남 2008.03.24
                .AddItem "주말법대" & Space(30) & "04"
                .AddItem "주말의대" & Space(30) & "05"
                
                .AddItem "야간법대" & Space(30) & "06"
                .AddItem "야간의대" & Space(30) & "07"
            
                .AddItem "선착순인문" & Space(30) & "11"
                .AddItem "선착순자연" & Space(30) & "12"
                
                .AddItem "선착순인문16" & Space(30) & "16"
                .AddItem "선착순자연17" & Space(30) & "17"
                
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
            
            .ListIndex = 0
        End With
        
            
        With cboPass(0)
            .Clear
            .AddItem "없음" & Space(30) & "X"
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(1)
            .Clear
            .AddItem "없음" & Space(30) & "X"
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(2)
            .Clear
            .AddItem "없음" & Space(30) & "X"
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(3)
            .Clear
            .AddItem "없음" & Space(30) & "X"
            .AddItem "노량진" & Space(30) & "N"
            .AddItem "강남" & Space(30) & "K"
            .AddItem "송파" & Space(30) & "S"
            .AddItem "송파 M" & Space(30) & "P"
            .AddItem "강남 M" & Space(30) & "M"
            
            .AddItem "주말법의대" & Space(30) & "W"
            .AddItem "야간법의대" & Space(30) & "Q"
            
            .AddItem "양재" & Space(30) & "J"
            .AddItem "부산" & Space(30) & "B"
            .AddItem "강남기숙(이천)" & Space(30) & "E"
            .ListIndex = 0
        End With
            
        sprPass.Tag = "0"           '<< 여러개 선택을 위해서
        
        sini_Path = App.Path & "\DAESUNG.INI"
        
        '>> 프로그램 INI 파일
        If Dir(sini_Path) = "" Then                                     '<< 파일이 없으면 생성
            sSort = insert_Order_ini_File("0,6/1,5/2,2/3,3/4,1/5,4/")
        End If
        
        sGbn = "STD020"
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SORT", "", sData, 255, sini_Path)         '>> SORT 순서
            sSort = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSort = insert_Order_ini_File("0,6/1,5/2,2/3,3/4,1/5,4/6,3/7,3/8,3/9,1/")
                sSort = "0,6/1,5/2,2/3,3/4,1/5,4/6,3/7,3/8,3/9,1/"
            End If
            
        Call init_Form(sSort)
        
    Me.Tag = ""
End Sub

Private Sub init_Form(ByVal aSort As String)
    Dim ni      As Integer
    Dim sDivs() As String
    Dim sDivC() As String
    
    fpTotCnt.value = 0
    
    optPassN.value = True
    optPassY.value = False
    
    fpTotS.value = 0
    fpTotE.value = 0
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""
    
    fpKor.value = 0
    fpMat.value = 0
    fpEng.value = 0
    
    sprPass.MaxRows = 0
    
    sDivs() = Split(aSort, "/", -1, vbTextCompare)
    For ni = 0 To UBound(sDivs) - 1 Step 1
        sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
        
        fpSort(CInt(sDivC(0))).value = CInt(sDivC(1))
    Next ni
    
    
'    fpSort(0).Value = 6         '<< 학생
'    fpSort(1).Value = 5         '<< 성명
'    fpSort(2).Value = 2         '<< 유.무
'    fpSort(3).Value = 3         '<< 합계
'    fpSort(4).Value = 1         '<< 계열
'    fpSort(5).Value = 4         '<< 선택

'    fpSort(6).Value = 3         '<< 언어
'    fpSort(7).Value = 3         '<< 수리
'    fpSort(8).Value = 3         '<< 외국어
'    fpSort(9).Value = 1         '<< 내신등급

End Sub

Private Function insert_Order_ini_File(ByVal aSort As String) As String
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sReturn     As String
    
    sGbn = "STD020"
         sReturn = basModule.WritePrivateProfileString(sGbn, "SORT", aSort, sini_Path)                 '<< SORT 순서
    
    insert_Order_ini_File = sReturn
    
End Function


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

    With sprPass
        For ni = 1 To 10 Step 1
            For nj = 0 To 9 Step 1
                If fpSort(nj).value = ni Then
                    nC = nC + 1
                    
                    Select Case nj
                        Case 0                      '<< 수험번호
                            .SortKey(nC) = 3
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "0," & CInt(Trim(fpSort(0).value)) & "/"
                        Case 1                      '<< 성명
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "1," & CInt(Trim(fpSort(1).value)) & "/"
                        Case 2                      '<< 유.무시험
                            .SortKey(nC) = 6
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "2," & CInt(Trim(fpSort(2).value)) & "/"
                        Case 3                      '<< 합계
                            .SortKey(nC) = 11
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "3," & CInt(Trim(fpSort(3).value)) & "/"
                        Case 4                      '<< 계열
                            .SortKey(nC) = 13
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "4," & CInt(Trim(fpSort(4).value)) & "/"
                        Case 5                      '<< 선택
                            .SortKey(nC) = 14
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "5," & CInt(Trim(fpSort(5).value)) & "/"
                            
                            
                        Case 6                      '<< 언어
                            .SortKey(nC) = 8
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "6," & CInt(Trim(fpSort(6).value)) & "/"
                        Case 7                      '<< 수리
                            .SortKey(nC) = 9
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "7," & CInt(Trim(fpSort(7).value)) & "/"
                        Case 8                      '<< 외국어
                            .SortKey(nC) = 10
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "8," & CInt(Trim(fpSort(8).value)) & "/"
                        Case 9                      '<< 내신등급
                            .SortKey(nC) = 12
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "9," & CInt(Trim(fpSort(9).value)) & "/"
                            
                    End Select
                    
                End If
            Next nj
        Next ni
        .SortKey(1) = 1
        .SortKey(2) = 3
        
        .Sort -1, -1, -1, -1, SortByRow
        
        sR = insert_Order_ini_File(sSort)
        
        sDivs() = Split(sR, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
            
            fpSort(CInt(sDivC(0))).value = CInt(sDivC(1))
        Next ni
    
    End With
    
End Sub










'>> 조회조건의 학생검색
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
    
    On Error GoTo ErrStmt
    
    cmdFind.Enabled = False
    
    chkAll.value = 0
    sprPass.MaxRows = 0
    fpTotCnt.value = 0
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, ACID, EXMID, STDNM, Birth_ymd_F, Birth_ymd, "
    sStr = sStr & "         EXMTYPE, EXMTYPE_NM,"
    sStr = sStr & "         K_NUM, M_NUM, E_NUM, TOT_NUM, N_NUM,"
    
    'sStr = sStr & "         GAEYUL, "
    sStr = sStr & "         SEL2_SCH, "                                     '< 2008.01.11 : 송파 M -> 제2지망
    
    sStr = sStr & "         KAEYOL_CD, KAEYOL_NM, "
    sStr = sStr & "         PASS1, PASS2, PASS3, PASS4 "
    sStr = sStr & "    FROM ("
            sStr = sStr & "  SELECT SCHNO, ACID, EXMID, STDNM, SUBSTR(Birth_ymd,1,4)||'-'||SUBSTR(Birth_ymd,5,2)  ||'-'||SUBSTR(Birth_ymd,7,2) AS Birth_ymd_F, Birth_ymd ,"
            sStr = sStr & "         EXMTYPE, DECODE(EXMTYPE,'0','무시험','1','유시험') AS EXMTYPE_NM,"
            sStr = sStr & "         K_NUM, E_NUM, M_NUM, N_NUM,"
            sStr = sStr & "         NVL( NVL(K_NUM,0)+NVL(E_NUM,0)+NVL(M_NUM,0), 0) AS  TOT_NUM,"
            sStr = sStr & "         SEL1, SEL3, "
            
            
'            sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
'            sStr = sStr & "             '사탐'"
'            sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
'            sStr = sStr & "             '과탐'"
'            sStr = sStr & "         END END GAEYUL,"
            sStr = sStr & "         SEL2_SCH, "                             '< 2008.01.11 : 송파 M -> 제2지망
            
            sStr = sStr & "         KAEYOL AS KAEYOL_CD,"
            
            '<< 계열 >> : 2008.01.09
            If Trim(basModule.SchCD) = "N" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                sStr = sStr & "                   '03','예체',"
                sStr = sStr & "                   '04','수리(나)',"
                sStr = sStr & "                   '05','인문수능',"
                sStr = sStr & "                   '06','자연수능',"
                
                sStr = sStr & "                   '07','신설인문',"
                sStr = sStr & "                   '08','신설자연',"
                sStr = sStr & "                   '09','신설수능인문',"
                sStr = sStr & "                   '10','신설수능자연',"
                
                sStr = sStr & "                   '11','편)인문',"
                sStr = sStr & "                   '12','편)자연',"
                sStr = sStr & "                   '13','편)예체',"
                sStr = sStr & "                   '14','편)수리(나)',"
                sStr = sStr & "                   '15','편)인문수능',"
                sStr = sStr & "                   '16','편)자연수능'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< 계열 >> : 2008.01.10
            ElseIf Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                
                sStr = sStr & "                   '04','주말법대',"
                sStr = sStr & "                   '05','주말의대',"
                sStr = sStr & "                   '06','야간법대',"
                sStr = sStr & "                   '07','야간의대',"
                
                sStr = sStr & "                   '11','선착순인문',"
                sStr = sStr & "                   '12','선착순자연',"
                
                sStr = sStr & "                   '16','선착순인문16',"
                sStr = sStr & "                   '17','선착순자연17'"
                
                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< 계열 >> : 2008.02.15
            ElseIf Trim(basModule.SchCD) = "S" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                sStr = sStr & "                   '03','예체능',"
                
                sStr = sStr & "                   '05','수능인문',"
                sStr = sStr & "                   '06','수능자연',"
                
                sStr = sStr & "                   '11','신설인문',"
                sStr = sStr & "                   '12','신설자연',"
                
                sStr = sStr & "                   '18','인문프리미엄',"
                sStr = sStr & "                   '19','자연프리미엄',"
                
                sStr = sStr & "                   '21','서울대특별인문',"
                sStr = sStr & "                   '22','서울대특별자연',"
                
                sStr = sStr & "                   '23','야간서울대인문',"
                sStr = sStr & "                   '24','야간서울대자연'"
                

                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< 계열 >> : 2008.02.15
            ElseIf Trim(basModule.SchCD) = "P" Then         '< 마송
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                sStr = sStr & "                   '03','특별인문',"
                sStr = sStr & "                   '04','특별자연'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            ElseIf Trim(basModule.SchCD) = "J" Then         '< 양재
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                sStr = sStr & "                   '11','신설인문',"
                sStr = sStr & "                   '12','신설자연',"
                
                sStr = sStr & "                   '18','인문프리미엄',"
                sStr = sStr & "                   '19','자연프리미엄'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            ElseIf Trim(basModule.SchCD) = "B" Then         '< 부산 : 2009.01.09
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연',"
                sStr = sStr & "                   '05','특별인문',"
                sStr = sStr & "                   '06','특별자연',"
                sStr = sStr & "                   '07','연고대인문',"
                sStr = sStr & "                   '08','연고대자연',"
                sStr = sStr & "                   '09','심화인문',"
                sStr = sStr & "                   '10','심화자연'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            Else
                sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
                sStr = sStr & "                   '02','자연'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
            End If
            
            sStr = sStr & "         PASS1, PASS2, PASS3, PASS4 "
            sStr = sStr & "    From CLSTD01TB"
            sStr = sStr & "   WHERE ACID  = '" & Trim(Right(cboHakwon.Text, 30)) & "'"
            sStr = sStr & "     AND EXMID > ' ' "           '> 결재한 학생
            sStr = sStr & "     AND CL_CLOSE IS NULL "      '> 완료여부 : 저장되면 YYMM값이 들어감.
            
            sStr = sStr & "     AND BIGO2 IS NULL"          '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    
    
    Select Case basModule.SchCD
        Case "K"
            sStr = sStr & "     AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            
        Case Else
            If optPassN.value = True Then
                'If Trim(basModule.SchCD) = "N" Then
                    'sStr = sStr & "     AND BIGO1 = '7' "                                      '> 완료여부 : 저장되면 YYMM값이 들어감.
                    'sStr = sStr & "     AND REGDATE <= TO_DATE('20080315','YYYYMMDD')"         '< 주의깊게 볼 것 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    
                    'sStr = sStr & "     AND BIGO1 = (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB WHERE ACID = '" & Trim(basModule.SchCD) & "') "      '> 완료여부 : 저장되면 YYMM값이 들어감. : 2008.02.28
                    sStr = sStr & "     AND BIGO1 = (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB ) "      '> 완료여부 : 저장되면 YYMM값이 들어감. : 2008.02.28
                'End If
            End If
            
    End Select
       
    sStr = sStr & "          )"
    sStr = sStr & "   WHERE SCHNO > ' ' "
    
'>> 합격처리할 학생 & 합격된 학생
    If optPassN.value = True Then
        sStr = sStr & " AND (PASS1 IS NULL AND PASS2 IS NULL AND PASS3 IS NULL AND PASS4 IS NULL)"
    ElseIf optPassY.value = True Then
        sStr = sStr & " AND (PASS1 > ' ' OR PASS2 > ' ' OR PASS3 > ' ' OR PASS4 > ' ') "
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
    
'>> 제2지망
    Select Case Trim(Right(cboSel2_Sch.Text, 30))
        Case "X"
            ' no action
        Case Else
            sStr = sStr & " AND SEL2_SCH = '" & Trim(Right(cboSel2_Sch.Text, 30)) & "'"
    End Select
    
'>> 계열
'    Select Case Trim(Right(cboKaeyol, 30))
'        Case "XX"
'            ' no action
'        Case "01", "03", "05"
'            sStr = sStr & "AND SEL1 > ' ' "
'        Case "02", "04", "06"
'            sStr = sStr & "AND SEL3 > ' ' "
'    End Select
    
    If Trim(Right(cboKaeyol.Text, 30)) = "ALL" Then
        'NO Action
    Else
        sStr = sStr & "    AND KAEYOL_CD = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    End If
    
'>> 수험번호
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = "" Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) = "" Then
        ' no action
    End If
       
'>> 학생명
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & " AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
    End If
'>> 주민번호
    If Trim(fpBirth_ymd.UnFmtText) > " " Then
        sStr = sStr & " AND Birth_ymd LIKE '" & Trim(fpBirth_ymd.UnFmtText) & "%'"
    End If
    
'>> 합계
    If fpTotS.value > 0 And fpTotE.value > 0 Then
        sStr = sStr & " AND ( TOT_NUM >= " & Trim(CStr(fpTotS.value)) & " AND TOT_NUM <= " & Trim(CStr(fpTotE.value)) & ")"
        
    ElseIf fpTotS.value > 0 And fpTotE.value = 0 Then
        sStr = sStr & " AND ( TOT_NUM >= " & Trim(CStr(fpTotS.value)) & " AND TOT_NUM <= 9999 )"
        
    ElseIf fpTotS.value = 0 And fpTotE.value > 0 Then
        sStr = sStr & " AND ( TOT_NUM >= 0 AND TOT_NUM <= " & Trim(CStr(fpTotE.value)) & ")"
        
    Else
        ' no action
    End If

    Select Case Trim(Right(cboExmType.Text, 30))
        Case "0"
            '>> 언어
                If fpKor.value > 0 Then
                    sStr = sStr & " AND K_NUM <= " & Trim(CStr(fpKor.value))
                End If
            '>> 수리
                If fpMat.value > 0 Then
                    sStr = sStr & " AND M_NUM <= " & Trim(CStr(fpMat.value))
                End If
            '>> 외국어
                If fpEng.value > 0 Then
                    sStr = sStr & " AND E_NUM <= " & Trim(CStr(fpEng.value))
                End If
        Case "1"
            '>> 언어
                If fpKor.value > 0 Then
                    sStr = sStr & " AND K_NUM >= " & Trim(CStr(fpKor.value))
                End If
            '>> 수리
                If fpMat.value > 0 Then
                    sStr = sStr & " AND M_NUM >= " & Trim(CStr(fpMat.value))
                End If
            '>> 외국어
                If fpEng.value > 0 Then
                    sStr = sStr & " AND E_NUM >= " & Trim(CStr(fpEng.value))
                End If
    End Select
    Text1.Text = sStr
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    

    
'    '>> 분원
'        sTmp = Trim(Right(cboHakwon.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> 수험번호
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
'    '>> 합계
'        If fpTot.Value > 0 Then
'            nTmp = CLng(fpTot.Value)
'            Set DBParam = DBCmd.CreateParameter("TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'        End If
'    '>> 학생명
'        If Trim(txtStdNM.Text) > " " Then
'            sTmp = "%" & Trim(txtStdNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'    '>> 주민번호
'        If Trim(fpBirth_ymd.UnFmtText) > " " Then
'            sTmp = "%" & Trim(fpBirth_ymd.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
            
                fpTotCnt.value = fpTotCnt.value + 1
            
                sprPass.MaxRows = sprPass.MaxRows + 1
                sprPass.Row = sprPass.MaxRows
                
                sprPass.Col = 1
                    sTmp = " ": If IsNull(.Fields("ACID")) = False Then sTmp = Trim(.Fields("ACID"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprPass.Col = sprPass.Col + 1       ' 수험번호
                    sTmp = " ": If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("Birth_ymd")) = False Then sTmp = Trim(.Fields("Birth_ymd"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("EXMTYPE")) = False Then sTmp = Trim(.Fields("EXMTYPE"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("EXMTYPE_NM")) = False Then sTmp = Trim(.Fields("EXMTYPE_NM"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprPass.Col = sprPass.Col + 1
                    nTmp = 0:     If IsNull(.Fields("K_NUM")) = False Then nTmp = CDbl(Trim(.Fields("K_NUM")))
                        Call basFunction.Set_SprType_Numeric(sprPass, 1, 1, 9999, "", nTmp)
                sprPass.Col = sprPass.Col + 1
                    nTmp = 0:   If IsNull(.Fields("M_NUM")) = False Then nTmp = CDbl(Trim(.Fields("M_NUM")))
                        Call basFunction.Set_SprType_Numeric(sprPass, 1, 1, 9999, "", nTmp)
                sprPass.Col = sprPass.Col + 1
                    nTmp = 0:   If IsNull(.Fields("E_NUM")) = False Then nTmp = CDbl(Trim(.Fields("E_NUM")))
                        Call basFunction.Set_SprType_Numeric(sprPass, 1, 1, 9999, "", nTmp)
                sprPass.Col = sprPass.Col + 1
                    nTmp = 0:   If IsNull(.Fields("TOT_NUM")) = False Then nTmp = CDbl(Trim(.Fields("TOT_NUM")))
                        Call basFunction.Set_SprType_Numeric(sprPass, 1, 1, 9999, "", nTmp)
                sprPass.Col = sprPass.Col + 1
                    nTmp = 0:   If IsNull(.Fields("N_NUM")) = False Then nTmp = CDbl(Trim(.Fields("N_NUM"))) '내신등급
                        Call basFunction.Set_SprType_Numeric(sprPass, 1, 1, 9999, "", nTmp)
                
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("KAEYOL_CD")) = False Then sTmp = Trim(.Fields("KAEYOL_CD"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("KAEYOL_NM")) = False Then sTmp = Trim(.Fields("KAEYOL_NM"))
                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
'                sprPass.Col = sprPass.Col + 1
'                    sTmp = " ": If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
'                        Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprPass.Col = sprPass.Col + 1
                    sTmp = " ": If IsNull(.Fields("SEL2_SCH")) = False Then sTmp = Trim(.Fields("SEL2_SCH"))        '< 2008.01.11 : 송파 M -> 제2지망
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "노량진"
                        Case "K"
                            sTmp = "강남"
                        Case "S"
                            sTmp = "송파"
                        Case "P"
                            sTmp = "송파 M"
                        Case "M"
                            sTmp = "강남 M"
                            
                        Case "W"
                            sTmp = "주말법의대"
                        Case "Q"
                            sTmp = "야간법의대"
                            
                        Case "J"
                            sTmp = "양재"
                        Case "B"
                            sTmp = "부산"
                       Case "E"
                            sTmp = "강남기숙(이천)"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                
                If IsNull(.Fields("PASS1")) = True Then
                    sprPass.Col = sprPass.Col + 1
                Else
                    sprPass.Col = sprPass.Col + 1
                    sTmp = Trim(.Fields("PASS1"))
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "노량진" & Space(30) & "N"
                        Case "K"
                            sTmp = "강남" & Space(30) & "K"
                        Case "S"
                            sTmp = "송파" & Space(30) & "S"
                        Case "P"
                            sTmp = "송파 M" & Space(30) & "P"
                        Case "M"
                            sTmp = "강남 M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "주말법의대" & Space(30) & "W"
                        Case "Q"
                            sTmp = "야간법의대" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "양재" & Space(30) & "J"
                        Case "B"
                            sTmp = "부산" & Space(30) & "B"
                        Case "E"
                            sTmp = "강남기숙(이천)" & Space(30) & "E"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", 300, sTmp)
                End If
                If IsNull(.Fields("PASS2")) = True Then
                    sprPass.Col = sprPass.Col + 1
                Else
                    sprPass.Col = sprPass.Col + 1
                    sTmp = Trim(.Fields("PASS2"))
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "노량진" & Space(30) & "N"
                        Case "K"
                            sTmp = "강남" & Space(30) & "K"
                        Case "S"
                            sTmp = "송파" & Space(30) & "S"
                        Case "P"
                            sTmp = "송파 M" & Space(30) & "P"
                        Case "M"
                            sTmp = "강남 M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "주말법의대" & Space(30) & "W"
                        Case "Q"
                            sTmp = "야간법의대" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "양재" & Space(30) & "J"
                        Case "B"
                            sTmp = "부산" & Space(30) & "B"
                         Case "E"
                            sTmp = "강남기숙(이천)" & Space(30) & "E"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", 300, sTmp)
                End If
                If IsNull(.Fields("PASS3")) = True Then
                    sprPass.Col = sprPass.Col + 1
                Else
                    sprPass.Col = sprPass.Col + 1
                    sTmp = Trim(.Fields("PASS3"))
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "노량진" & Space(30) & "N"
                        Case "K"
                            sTmp = "강남" & Space(30) & "K"
                        Case "S"
                            sTmp = "송파" & Space(30) & "S"
                        Case "P"
                            sTmp = "송파 M" & Space(30) & "P"
                        Case "M"
                            sTmp = "강남 M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "주말법의대" & Space(30) & "W"
                        Case "Q"
                            sTmp = "야간법의대" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "양재" & Space(30) & "J"
                        Case "B"
                            sTmp = "부산" & Space(30) & "B"
                        Case "E"
                            sTmp = "강남기숙(이천)" & Space(30) & "E"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", 300, sTmp)
                End If
                If IsNull(.Fields("PASS4")) = True Then
                    sprPass.Col = sprPass.Col + 1
                Else
                    sprPass.Col = sprPass.Col + 1
                    sTmp = Trim(.Fields("PASS4"))
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "노량진" & Space(30) & "N"
                        Case "K"
                            sTmp = "강남" & Space(30) & "K"
                        Case "S"
                            sTmp = "송파" & Space(30) & "S"
                        Case "P"
                            sTmp = "송파 M" & Space(30) & "P"
                        Case "M"
                            sTmp = "강남 M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "주말법의대" & Space(30) & "W"
                        Case "Q"
                            sTmp = "야간법의대" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "양재" & Space(30) & "J"
                        Case "B"
                            sTmp = "부산" & Space(30) & "B"
                        Case "E"
                            sTmp = "강남기숙(이천)" & Space(30) & "E"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", 300, sTmp)
                End If
                
                sprPass.Col = sprPass.Col + 1:  Call basFunction.Set_SprType_ChkBox(sprPass)
                
                .MoveNext
            Next nRec
            
            sprPass.Row = 1:       sprPass.Row2 = sprPass.MaxRows
            sprPass.Col = 1:       sprPass.Col2 = sprPass.MaxCols
            sprPass.BlockMode = True
                '.BackColor = basModule.BackColor2
                sprPass.BackColor = &HFFFFFF
                sprPass.BackColorStyle = BackColorStyleUnderGrid
                
                sprPass.Lock = True
                sprPass.Protect = True
            sprPass.BlockMode = False
            
            If sprPass.MaxRows > 0 Then
                Call cmdSort_Click
            End If
            
        End If
    End With
    
    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "학생 합격 및 확인"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    sprPass.SetFocus
    'sprPass.SetActiveCell 1, 1
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    On Error GoTo 0
    MsgBox "합격처리 및 확인 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생 합격 및 확인"
End Sub
















'>> 선택 ## multi 선택
Private Sub sprPass_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    
    If Row < 1 Then Exit Sub

    With sprPass
        If .MaxRows < 1 Then Exit Sub

        sprPass.Enabled = False
        
            If .Tag = "0" Then
                .Row = CLng(.Tag):      .Row2 = .Row
                .Col = 1:               .Col2 = .MaxCols
                .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = &HFFFFFF
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = Row
                    .Col = .MaxCols
                    .value = 0
                
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
                
                .Col = .MaxCols:    .value = 1
                
                .Tag = Trim(CStr(Row))
            ElseIf .Tag > "0" Then
                .Row = Row
                .Col = .MaxCols
                If .value = 1 Then
                    .value = 0
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = &HFFFFFF
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                Else
                    .value = 1
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = basModule.SelectColor2
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                End If
            
            End If
            
        sprPass.Enabled = True

        sprPass.SetFocus
        'sprPass.SetActiveCell Col, Row

    End With
End Sub

Private Sub sprPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nS      As Long
    Dim nE      As Long
    
    Dim nRow    As Long
    
    With sprPass
    
        If .MaxRows = 0 Then Exit Sub
        
        Select Case Shift
'            Case 0
'                Call sprPass_Click(1, .ActiveRow)
                
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
    
    With sprPass
        If .MaxRows = 0 Then Exit Sub
            
        If chkAll.value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = &HFFFFFF
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .value = 1
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









































'======================================================== 합격학생 등록하기 ============================================================

'>> 합격학원 넣기
Private Sub cboPass_Click(Index As Integer)
    Dim ni      As Long
    Dim nRec    As Long
    
    Dim sTmp    As String
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    With sprPass
        If .MaxRows = 0 Then Exit Sub
        
        nRec = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = .MaxCols
            If .value = 1 Then
                nRec = 1
                Exit For
            End If
        Next ni
        
        If nRec = 0 Then
            MsgBox "합격학생과 합격학원을 선택하여 주십시요.", vbExclamation + vbOKOnly, "합격자 처리"
            Exit Sub
        End If
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = .MaxCols
            
            If .value = 1 Then
                Select Case Index
                    Case 0
                        If Trim(Right(cboPass(0).Text, 30)) <> "X" Then
                            .Col = 16
                            sTmp = Trim(cboPass(0).Text)
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        Else
                            .Col = 16
                            sTmp = ""
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        End If
                    Case 1
                        If Trim(Right(cboPass(1).Text, 30)) <> "X" Then
                            .Col = 17
                            sTmp = Trim(cboPass(1).Text)
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        Else
                            .Col = 17
                            sTmp = ""
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        End If
                    Case 2
                        If Trim(Right(cboPass(2).Text, 30)) <> "X" Then
                            .Col = 18
                            sTmp = Trim(cboPass(2).Text)
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        Else
                            .Col = 18
                            sTmp = ""
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        End If
                    Case 3
                        If Trim(Right(cboPass(3).Text, 30)) <> "X" Then
                            .Col = 19
                            sTmp = Trim(cboPass(3).Text)
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        Else
                            .Col = 19
                            sTmp = ""
                            Call basFunction.Set_SprType_Text(sprPass, "center", "left", 300, sTmp)
                        End If
                End Select
            End If
        Next ni
    End With
    
End Sub



'>> 합격생 등록하기
Private Sub cmdSave_Click()
    Dim bRet        As Boolean
    
    Dim ni      As Long
    Dim nRec    As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    With sprPass
        If .MaxRows = 0 Then Exit Sub
        
'>> 체크조건
        nRec = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = .MaxCols
            If .value = 1 Then
                .Col = 16
                If Trim(.Text) > " " Then
                    nRec = 1
                    Exit For
                End If
                
                .Col = 17
                If Trim(.Text) > " " Then
                    nRec = 1
                    Exit For
                End If
                
                .Col = 18
                If Trim(.Text) > " " Then
                    nRec = 1
                    Exit For
                End If
                
                .Col = 19
                If Trim(.Text) > " " Then
                    nRec = 1
                    Exit For
                End If
                
            End If
        Next ni
        
        If nRec = 0 Then
            MsgBox "합격학생을 선택하여 주십시요.", vbExclamation + vbOKOnly, "합격자 처리"
            Exit Sub
        End If
    End With
    
    On Error GoTo ErrStmt
    
    cmdSave.Enabled = False
        bRet = Save_STD_Data
        
    cmdSave.Enabled = True
    
    If bRet = True Then
        MsgBox "합격자 등록 완료하였습니다.", vbInformation + vbOKOnly, "합격자 처리"
    Else
        MsgBox "합격자 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "합격자 처리"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "합격자 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "합격자 처리"
    On Error GoTo 0
    
End Sub





'>> 학생코드가 유일하므로
'>> 합격자 처리를 학생코드로만 업데이트 합니다.
Private Function Save_STD_Data() As Boolean
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
    
    Dim nRec        As Long         '<< 처리해야 할 수
    Dim nTot        As Long         '<< 처리한 수
    
    bRet = False
    nRec = 0
    nTot = 0
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    For nRow = 1 To sprPass.MaxRows Step 1
        
        sprPass.Row = nRow
        sprPass.Col = sprPass.MaxCols
        
        If sprPass.value = 1 Then
        
            nRec = nRec + 1
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
        
            sStr = ""
            sStr = sStr & "  Update CLSTD01TB"
            sStr = sStr & "     SET PASS1 = ?, "
            sStr = sStr & "         PASS2 = ?, "
            sStr = sStr & "         PASS3 = ?, "
            sStr = sStr & "         PASS4 = ? "
            sStr = sStr & "   WHERE SCHNO = ? "
            
            '>> 1차 합격
                sprPass.Col = 16
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 2차 합격
                sprPass.Col = 17
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 3차 합격
                sprPass.Col = 18
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 4차 합격
                sprPass.Col = 19
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                    
            '>> 학생코드
                sprPass.Col = 2
                sTmp = Trim(sprPass.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
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
        Save_STD_Data = True
    Else
        Save_STD_Data = False
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_STD_Data = False
    
End Function





'>> 취소처리
Private Sub cmdCancel_Click()
    Dim nRow        As Long
    Dim nCnt        As Long
    
    Dim nChkRow     As Long
    
    Dim sAcID       As String
    Dim sSchNO      As String
    Dim sExmID      As String
    
    
    If optPassY.value = False Then
        MsgBox "합격생만 조회시에 가능합니다.", vbExclamation + vbOKOnly, "취소처리"
        Exit Sub
    End If
    
    If sprPass.MaxRows = 0 Then
        MsgBox "합격생만 조회후 처리하십시요.", vbExclamation + vbOKOnly, "취소처리"
        Exit Sub
    End If
    
    With sprPass
        nCnt = 0
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols
            
            If .value = 1 Then
            
                nCnt = nCnt + 1
                'nChkRow = .Row
                
                .Row = nRow
                
                .Col = 1:           sAcID = Trim(.Text)
                .Col = 2:           sSchNO = Trim(.Text)
                .Col = 3:           sExmID = Trim(.Text)
                
                Call Process_Cancel(sAcID, sSchNO, sExmID)
                
            End If
            
        Next nRow
    End With
    
    MsgBox "완료하였습니다.", vbInformation + vbOKOnly, "취소처리"
    
'    If nCnt > 1 Or nCnt = 0 Then
'        MsgBox "합격생 1명만 취소처리가능합니다.", vbExclamation + vbOKOnly, "취소처리"
'        Exit Sub
'    End If
    
'    With sprPass
'        .Row = nChkRow
'        .Col = 1:           sAcID = Trim(.Text)
'        .Col = 2:           sSchNO = Trim(.Text)
'        .Col = 3:           sExmID = Trim(.Text)
'
'        Call Process_Cancel(sAcID, sSchNO, sExmID)
'
'    End With
    
End Sub



Private Sub Process_Cancel(ByVal aAcID As String, ByVal aSchNO As String, ByVal aExmID As String)
    
    Dim bRet        As Boolean
    
    On Error GoTo ErrStmt
    
    '1 학원
    '2 학생
    '3 수험번호
    
    cmdCancel.Enabled = False
        bRet = Cancel_StdOut(aAcID, aSchNO, aExmID)
        
    cmdCancel.Enabled = True
    
    If bRet = True Then
        'MsgBox "학생 합격취소 하였습니다.", vbInformation + vbOKOnly, "학생 취소하기"
    Else
        MsgBox "학생 합격취소시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생 취소하기"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "학생 합격취소시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생 취소하기"
    On Error GoTo 0
    
End Sub

'>> 합격취소하기
Private Function Cancel_StdOut(ByVal aAcID As String, ByVal aSchNO As String, ByVal aExmID As String) As Boolean
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

    
    
    sStr = ""
    sStr = sStr & " UPDATE CLSTD01TB "
    sStr = sStr & "    SET PASS1 = '', "
    sStr = sStr & "        PASS2 = '', "
    sStr = sStr & "        PASS3 = '', "
    sStr = sStr & "        PASS4 = '', "
    
    sStr = sStr & "        CY_ACNT = '', "
    sStr = sStr & "        TOT_AMT = 0 , "
    
    sStr = sStr & "        BASE_AMT1  = 0 , "
    sStr = sStr & "        BASE_AMT2  = 0 , "
    sStr = sStr & "        BASE_AMT3  = 0 , "
    sStr = sStr & "        BASE_AMT4  = 0 , "
    sStr = sStr & "        BASE_AMT5  = 0 , "
    sStr = sStr & "        BASE_AMT6  = 0 , "
    sStr = sStr & "        BASE_AMT7  = 0 , "
    sStr = sStr & "        BASE_AMT8  = 0 , "
    sStr = sStr & "        BASE_AMT9  = 0 , "
    sStr = sStr & "        BASE_AMT10 = 0 , "
    
    sStr = sStr & "        TAMGU_AMT1  = 0 , "
    sStr = sStr & "        TAMGU_AMT2  = 0 , "
    sStr = sStr & "        TAMGU_AMT3  = 0 , "
    sStr = sStr & "        TAMGU_AMT4  = 0 , "
    sStr = sStr & "        TAMGU_AMT5  = 0 , "
    sStr = sStr & "        TAMGU_AMT6  = 0 , "
    sStr = sStr & "        TAMGU_AMT7  = 0 , "
    sStr = sStr & "        TAMGU_AMT8  = 0 , "
    sStr = sStr & "        TAMGU_AMT9  = 0 , "
    sStr = sStr & "        TAMGU_AMT10 = 0 , "
    sStr = sStr & "        TAMGU_AMT11 = 0 , "
    sStr = sStr & "        TAMGU_AMT12 = 0  "
    
    sStr = sStr & "  WHERE ACID  = '" & aAcID & "'"
    sStr = sStr & "    AND SCHNO = '" & aSchNO & "'"
    
    nExe = 0
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nExe = 1 Then
        bRet = True
        basDataBase.DBConn.CommitTrans
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    Cancel_StdOut = bRet

    Set DBCmd = Nothing
    Set DBParam = Nothing

    
'    sStr = ""
'    sStr = sStr & " INSERT INTO CLSTD91TB"
'    sStr = sStr & " SELECT *"
'    sStr = sStr & "   FROM CLSTD01TB "
'    sStr = sStr & "   WHERE SCHNO   = '" & Trim(aSchNO) & "'"
'
'
'    DBCmd.CommandText = sStr
'    DBCmd.CommandType = adCmdText
'    DBCmd.CommandTimeout = 30
'
'    DBCmd.Execute nExe, , -1
'
'
'    Do While basDataBase.DBConn.State And adStateExecuting
'        DoEvents
'    Loop
'
'    If nExe = 1 Then
'        nExe = 0
'
'        '-----------------------------------------------------------------------------------------------------
'        Select Case Trim(basModule.SchCD)
'            Case "S"
'                sStr = ""
'                sStr = sStr & " UPDATE CLSTD01TB "
'                sStr = sStr & "    SET PASS1   = '',"
'                sStr = sStr & "        PASS2   = '',"
'                sStr = sStr & "        PASS3   = '',"
'                sStr = sStr & "        PASS4   = '',"
'                sStr = sStr & "        CY_ACNT = '',"
'                sStr = sStr & "        TOT_AMT = 0 "
'                sStr = sStr & "  WHERE SCHNO   = '" & Trim(aSchNO) & "'"
'            Case Else
'                sStr = ""
'                sStr = sStr & " DELETE "
'                sStr = sStr & "   FROM CLSTD01TB "
'                sStr = sStr & "  WHERE SCHNO   = '" & Trim(aSchNO) & "'"
'        End Select
'
'
'        DBCmd.CommandText = sStr
'        DBCmd.CommandType = adCmdText
'        DBCmd.CommandTimeout = 30
'
'        DBCmd.Execute nExe, , -1
'
'
'        Do While basDataBase.DBConn.State And adStateExecuting
'            DoEvents
'        Loop
'
'        If nExe = 1 Then
'            nExe = 0
'            On Error Resume Next
'
'            sStr = ""
'            sStr = sStr & " INSERT INTO CLSTD92TB (SCHNO, ACID, EXMID, TIMESTAMP) "
'            sStr = sStr & " VALUES( "
'            sStr = sStr & "         '" & Trim(aSchNO) & "',"
'            sStr = sStr & "         '" & Trim(aAcID) & "',"
'            sStr = sStr & "         '" & Trim(aExmID) & "',"
'            sStr = sStr & "         SYSDATE"
'            sStr = sStr & "       ) "
'
'            DBCmd.CommandText = sStr
'            DBCmd.CommandType = adCmdText
'            DBCmd.CommandTimeout = 30
'
'            DBCmd.Execute nExe, , -1
'
'
'            Do While basDataBase.DBConn.State And adStateExecuting
'                DoEvents
'            Loop
'
'            If nExe = 1 Then
'                bRet = True
'            Else
'                nExe = 0
'
'                On Error GoTo 0
'                On Error GoTo ErrStmt
'
'                sStr = ""
'                sStr = sStr & " UPDATE CLSTD92TB "
'                sStr = sStr & "    SET ACID  = '" & Trim(aSchNO) & "',"
'                sStr = sStr & "        EXMID = '" & Trim(aExmID) & "',"
'                sStr = sStr & "        TIMESTAMP = SYSDATE "
'                sStr = sStr & "  WHERE SCHNO = '" & Trim(aSchNO) & "'"
'
'                DBCmd.CommandText = sStr
'                DBCmd.CommandType = adCmdText
'                DBCmd.CommandTimeout = 30
'
'                DBCmd.Execute nExe, , -1
'
'                Do While basDataBase.DBConn.State And adStateExecuting
'                    DoEvents
'                Loop
'
'                If nExe = 1 Then
'                    bRet = True
'                End If
'
'            End If
'        End If
'        '-----------------------------------------------------------------------------------------------------
'
'    End If
'
'
'    Cancel_StdOut = bRet
'
'    Set DBCmd = Nothing
'    Set DBParam = Nothing
'
'    basDataBase.DBConn.CommitTrans
    
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Cancel_StdOut = bRet
End Function

Private Sub txtStdNM_Change()
'    If KeyCode = vbKeyReturn Then
'        If Trim(txtStdNM.Text) > " " Then
'            Call cmdFind_Click
'        End If
'    End If
End Sub
