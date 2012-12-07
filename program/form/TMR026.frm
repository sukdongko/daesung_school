VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR026 
   Caption         =   "시간표 만들기 >> 이동수업 시간표 등록 CP"
   ClientHeight    =   11190
   ClientLeft      =   570
   ClientTop       =   1800
   ClientWidth     =   17385
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11190
   ScaleWidth      =   17385
   Begin VB.Frame Frame9 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame9"
      Height          =   2685
      Left            =   30
      TabIndex        =   42
      Top             =   8490
      Width           =   17235
      Begin VB.Frame Frame8 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame8"
         Height          =   2625
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   17175
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '없음
            Caption         =   "Frame6"
            Height          =   375
            Left            =   5490
            TabIndex        =   44
            Top             =   0
            Width           =   11415
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
               Left            =   210
               TabIndex        =   15
               Top             =   0
               Width           =   645
            End
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   0
               Left            =   510
               TabIndex        =   16
               Top             =   120
               Visible         =   0   'False
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
               Left            =   1170
               TabIndex        =   17
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   1830
               TabIndex        =   18
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   2460
               TabIndex        =   19
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   3090
               TabIndex        =   20
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   3720
               TabIndex        =   21
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   4380
               TabIndex        =   22
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   5040
               TabIndex        =   23
               Top             =   30
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
               MaxValue        =   "16"
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
               Left            =   5670
               TabIndex        =   24
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   9
               Left            =   6300
               TabIndex        =   25
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   10
               Left            =   6960
               TabIndex        =   26
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   11
               Left            =   7590
               TabIndex        =   27
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   12
               Left            =   8250
               TabIndex        =   28
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   13
               Left            =   8880
               TabIndex        =   29
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   14
               Left            =   9510
               TabIndex        =   30
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   15
               Left            =   10170
               TabIndex        =   31
               Top             =   30
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
               MaxValue        =   "16"
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
               Index           =   16
               Left            =   10800
               TabIndex        =   32
               Top             =   30
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
               MaxValue        =   "16"
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
               TabIndex        =   45
               Top             =   75
               Width           =   645
            End
         End
         Begin FPSpread.vaSpread sprSTD 
            Height          =   2175
            Left            =   0
            TabIndex        =   14
            Top             =   420
            Width           =   17145
            _Version        =   393216
            _ExtentX        =   30242
            _ExtentY        =   3836
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
            MaxCols         =   28
            SpreadDesigner  =   "TMR026.frx":0000
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '투명
            Caption         =   "헤더를 누르면 SORT가 됩니다."
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   780
            TabIndex        =   48
            Top             =   120
            Width           =   2805
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame7"
      Height          =   6195
      Left            =   30
      TabIndex        =   36
      Top             =   2280
      Width           =   17235
      Begin VB.Frame Frame3 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame3"
         Height          =   6135
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   17175
         Begin VB.Frame Frame11 
            BackColor       =   &H00F7EFE7&
            Height          =   2745
            Left            =   14400
            TabIndex        =   47
            Top             =   660
            Width           =   2655
            Begin VB.ComboBox cboLsnTypeCP 
               Height          =   300
               Left            =   150
               Style           =   2  '드롭다운 목록
               TabIndex        =   11
               Top             =   330
               Width           =   975
            End
            Begin VB.CommandButton cmdBanToGwamok 
               Caption         =   "반별 과목 등록하기"
               Height          =   465
               Left            =   150
               TabIndex        =   13
               Top             =   1500
               Width           =   2355
            End
            Begin VB.CommandButton cmdSearchSaveData 
               Caption         =   "기 등록내용 가져오기"
               Height          =   465
               Left            =   150
               TabIndex        =   12
               Top             =   810
               Width           =   2355
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00F7EFE7&
            Height          =   675
            Left            =   2190
            TabIndex        =   46
            Top             =   0
            Width           =   3195
            Begin VB.CommandButton cmdinPut 
               Caption         =   "선택과목 등록"
               Height          =   435
               Left            =   1260
               TabIndex        =   5
               Top             =   150
               Width           =   1785
            End
            Begin VB.ComboBox cboLsnType 
               Height          =   300
               Left            =   240
               Style           =   2  '드롭다운 목록
               TabIndex        =   4
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '없음
            Caption         =   "Frame2"
            Height          =   2655
            Left            =   13020
            TabIndex        =   38
            Top             =   750
            Width           =   1215
            Begin VB.Frame Frame1 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '없음
               Caption         =   "Frame1"
               Height          =   2595
               Left            =   30
               TabIndex        =   39
               Top             =   30
               Width           =   1155
               Begin VB.OptionButton optLsn 
                  BackColor       =   &H008080FF&
                  Caption         =   "선택1"
                  Height          =   375
                  Index           =   0
                  Left            =   120
                  TabIndex        =   7
                  Top             =   210
                  Width           =   915
               End
               Begin VB.OptionButton optLsn 
                  BackColor       =   &H0000FFFF&
                  Caption         =   "선택2"
                  Height          =   375
                  Index           =   1
                  Left            =   120
                  TabIndex        =   8
                  Top             =   810
                  Width           =   915
               End
               Begin VB.OptionButton optLsn 
                  BackColor       =   &H0000FF00&
                  Caption         =   "선택3"
                  Height          =   375
                  Index           =   2
                  Left            =   120
                  TabIndex        =   9
                  Top             =   1440
                  Width           =   915
               End
               Begin VB.OptionButton optLsn 
                  BackColor       =   &H00FF8080&
                  Caption         =   "선택4"
                  Height          =   375
                  Index           =   3
                  Left            =   120
                  TabIndex        =   10
                  Top             =   2070
                  Width           =   915
               End
            End
         End
         Begin VB.CommandButton cmdGetLsn 
            Caption         =   "반 선택하기"
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
            Left            =   150
            TabIndex        =   3
            Top             =   210
            Width           =   1725
         End
         Begin FPSpread.vaSpread sprBanChk 
            Height          =   5355
            Left            =   30
            TabIndex        =   6
            Top             =   750
            Width           =   12915
            _Version        =   393216
            _ExtentX        =   22781
            _ExtentY        =   9446
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
            SpreadDesigner  =   "TMR026.frx":2009
         End
         Begin VB.Label lblStatus 
            BackStyle       =   0  '투명
            Caption         =   "상태"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   285
            Left            =   5490
            TabIndex        =   40
            Top             =   540
            Width           =   11415
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   $"TMR026.frx":3E1C
            Height          =   645
            Left            =   5520
            TabIndex        =   41
            Top             =   0
            Width           =   8715
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame5"
      Height          =   2235
      Left            =   30
      TabIndex        =   33
      Top             =   0
      Width           =   13755
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame4"
         Height          =   2175
         Left            =   30
         TabIndex        =   34
         Top             =   30
         Width           =   13695
         Begin VB.CommandButton cmdFind 
            Caption         =   "전체 반 선택과목 신청조회 (&F)"
            Height          =   375
            Left            =   2490
            TabIndex        =   1
            Top             =   30
            Width           =   3045
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   1290
            Style           =   2  '드롭다운 목록
            TabIndex        =   0
            Top             =   75
            Width           =   1005
         End
         Begin FPSpread.vaSpread sprTotGwamok 
            Height          =   1695
            Left            =   30
            TabIndex        =   2
            Top             =   435
            Width           =   13605
            _Version        =   393216
            _ExtentX        =   23998
            _ExtentY        =   2990
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
            MaxCols         =   21
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "TMR026.frx":3F06
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   630
            TabIndex        =   35
            Top             =   150
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "TMR026"
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

Private Const nRowHeight = 15

Private Sub Form_Unload(Cancel As Integer)
    Unload TMR028
    
End Sub


Private Sub Form_Load()
    Dim ni      As Integer
    
    Me.Move 0, 0, 17500, 11600
    
    Me.Tag = "LOAD"
        With cboKaeyol
            .Clear
            .AddItem "인문" & Space(30) & "01"
            .AddItem "자연" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With sprTotGwamok
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .Tag = "0"
            .MaxRows = 0
        End With
        
        With sprBanChk
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxRows = 0
            
            '< 이동반 등록 >
            Call Add_MV_Lsn
            
        End With
        
        
        With sprSTD
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxRows = 0
        End With
        
        With cboLsnType
            .Clear
            .AddItem "A type" & Space(30) & "A"
            .AddItem "B type" & Space(30) & "B"
            .AddItem "C type" & Space(30) & "C"
            
            .ListIndex = 0
        End With
        
        With cboLsnTypeCP
            .Clear
            .AddItem "A type" & Space(30) & "A"
            .AddItem "B type" & Space(30) & "B"
            .AddItem "C type" & Space(30) & "C"
            
            .ListIndex = 0
        End With
        
        
        lblStatus.Caption = ""
        
        optLsn(0).value = True
        optLsn(1).value = False
        optLsn(2).value = False
        optLsn(3).value = False
        
        For ni = 1 To 16 Step 1
            fpSort(ni).value = ni
        Next ni
        
    Me.Tag = ""
    
End Sub

Private Sub cboLsnType_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    cboLsnTypeCP.ListIndex = cboLsnType.ListIndex
    
End Sub

Private Sub cboLsnTypeCP_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    cboLsnType.ListIndex = cboLsnTypeCP.ListIndex
    
End Sub



'## 이동반 생성
Private Sub Add_MV_Lsn()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Long
    
    sprBanChk.MaxRows = 0
    
    On Error Resume Next
    
    sStr = ""
    sStr = sStr & "    SELECT LSNCD, LSNNM, LSNCDNM, LSNCAPA, SEL_OK, LSN_CL, 0 AS S_LSN"
    sStr = sStr & "      From SDLSN02TB"
    sStr = sStr & "     WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "       AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "     ORDER BY LSNCDNM"
    
    
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
            
            '< 합 계 >
            sprBanChk.MaxRows = sprBanChk.MaxRows + 1
                sprBanChk.Row = sprBanChk.MaxRows:    sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight


            For nRec = 1 To .RecordCount Step 1
                sprBanChk.MaxRows = sprBanChk.MaxRows + 1
                sprBanChk.Row = sprBanChk.MaxRows:    sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight


                sprBanChk.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprBanChk.Col = sprBanChk.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprBanChk.Col = sprBanChk.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> 총인원
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
                '<< 인문자연 공통 : 8 과목 >>
                For nCol = 1 To 8 Step 1
                    sprBanChk.Col = sprBanChk.Col + 1
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                Next nCol
                
                '사탐/ 과탐 공통
                For nCol = 9 To 11 Step 1
                    sprBanChk.Col = sprBanChk.Col + 1
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                Next nCol
                        
                
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 제2선택
                sprBanChk.Col = sprBanChk.Col + 1
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 언
                sprBanChk.Col = sprBanChk.Col + 1
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    
                '> 수
                sprBanChk.Col = sprBanChk.Col + 1
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    
                '> 사
                sprBanChk.Col = sprBanChk.Col + 1
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    
                '> 탐
                sprBanChk.Col = sprBanChk.Col + 1
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                .MoveNext       '<< 다음항목
                
            Next nRec
        End If
        
        With sprBanChk
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .ColsFrozen = 4
            
        '>> spread lock
            .Row = 1:       .Row2 = 1
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
End Sub



'## 계열선택
Private Sub cboKaeyol_Click()
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"
            With sprTotGwamok
                .Row = SpreadHeader + 1
                .Col = 5:           .Text = "국사"
                .Col = .Col + 1:    .Text = "윤리"
                .Col = .Col + 1:    .Text = "경제"
                .Col = .Col + 1:    .Text = "한근"
                .Col = .Col + 1:    .Text = "세계사"
                
                .Col = .Col + 1:    .Text = "경지"
                .Col = .Col + 1:    .Text = "한지"
                .Col = .Col + 1:    .Text = "정치"
                .Col = .Col + 1:    .Text = "사문"
                .Col = .Col + 1:    .Text = "법사"
                
                .Col = .Col + 1:    .Text = "세지"
                
                .MaxRows = 0
                
            End With
            
            With sprBanChk
                .Row = SpreadHeader + 1
                .Col = 5:           .Text = "국사"
                .Col = .Col + 1:    .Text = "윤리"
                .Col = .Col + 1:    .Text = "경제"
                .Col = .Col + 1:    .Text = "한근"
                .Col = .Col + 1:    .Text = "세계사"
                
                .Col = .Col + 1:    .Text = "경지"
                .Col = .Col + 1:    .Text = "한지"
                .Col = .Col + 1:    .Text = "정치"
                .Col = .Col + 1:    .Text = "사문"
                .Col = .Col + 1:    .Text = "법사"
                
                .Col = .Col + 1:    .Text = "세지"
                
                .MaxRows = 0
                
            End With
            
            With sprSTD
                .Row = SpreadHeader + 1
                .Col = 13:          .Text = "국사"
                .Col = .Col + 1:    .Text = "윤리"
                .Col = .Col + 1:    .Text = "경제"
                .Col = .Col + 1:    .Text = "한근"
                .Col = .Col + 1:    .Text = "세계사"
                
                .Col = .Col + 1:    .Text = "경지"
                .Col = .Col + 1:    .Text = "한지"
                .Col = .Col + 1:    .Text = "정치"
                .Col = .Col + 1:    .Text = "사문"
                .Col = .Col + 1:    .Text = "법사"
                
                .Col = .Col + 1:    .Text = "세지"
                
                .MaxRows = 0
                
            End With
            
            '< 이동반 등록 >
            Call Add_MV_Lsn
            
        Case "02"
            With sprTotGwamok
                .Row = SpreadHeader + 1
                .Col = 5:           .Text = "물1"
                .Col = .Col + 1:    .Text = "화1"
                .Col = .Col + 1:    .Text = "생1"
                .Col = .Col + 1:    .Text = "지1"
                .Col = .Col + 1:    .Text = "물2"
                
                .Col = .Col + 1:    .Text = "화2"
                .Col = .Col + 1:    .Text = "생2"
                .Col = .Col + 1:    .Text = "지2"
                .Col = .Col + 1:    .Text = "-"
                .Col = .Col + 1:    .Text = "-"
                
                .Col = .Col + 1:    .Text = "-"
                
                .MaxRows = 0
                
            End With
            
            With sprBanChk
                .Row = SpreadHeader + 1
                .Col = 5:           .Text = "물1"
                .Col = .Col + 1:    .Text = "화1"
                .Col = .Col + 1:    .Text = "생1"
                .Col = .Col + 1:    .Text = "지1"
                .Col = .Col + 1:    .Text = "물2"
                
                .Col = .Col + 1:    .Text = "화2"
                .Col = .Col + 1:    .Text = "생2"
                .Col = .Col + 1:    .Text = "지2"
                .Col = .Col + 1:    .Text = "-"
                .Col = .Col + 1:    .Text = "-"
                
                .Col = .Col + 1:    .Text = "-"
                
                .MaxRows = 0
                
            End With
            
            With sprSTD
                .Row = SpreadHeader + 1
                .Col = 13:          .Text = "물1"
                .Col = .Col + 1:    .Text = "화1"
                .Col = .Col + 1:    .Text = "생1"
                .Col = .Col + 1:    .Text = "지1"
                .Col = .Col + 1:    .Text = "물2"
                
                .Col = .Col + 1:    .Text = "화2"
                .Col = .Col + 1:    .Text = "생2"
                .Col = .Col + 1:    .Text = "지2"
                .Col = .Col + 1:    .Text = "-"
                .Col = .Col + 1:    .Text = "-"
                
                .Col = .Col + 1:    .Text = "-"
                
                .MaxRows = 0
                
            End With
            
            '< 이동반 등록 >
            Call Add_MV_Lsn
            
    End Select
End Sub





Private Sub cmdFind_Click()
    cmdFind.Enabled = False
    
    Call Fill_sprTotGwamok                  '< 과목내역
    cmdFind.Enabled = True
    
End Sub

Private Sub Exec_sprTotGwamok_Formula()
    Dim nCol        As Long
    
    With sprTotGwamok

     '>> 열 합계 -------------------------------------------------------
            For nCol = 4 To (.MaxCols - 1) Step 1
                .Row = 1
                .Col = nCol
                
                .CellType = CellTypeNumber
                .TypeVAlign = TypeVAlignCenter
                .TypeNumberDecPlaces = 0
                .TypeNumberMin = -9999
                .TypeNumberMax = 9999
                
                .TypeNumberShowSep = False
            Next nCol
            
            For nCol = 4 To (.MaxCols - 1) Step 1               '<<
                .Row = 1
                .Col = nCol
                .FormulaSync = True
                .Formula = "SUM(#2:#" & Trim(CStr(.MaxRows)) & ")"
                
            Next nCol
            
    End With
    
End Sub



'## 전체 과목별 학생수
Private Sub Fill_sprTotGwamok()
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
    
    On Error GoTo ErrStmt
    
    sprTotGwamok.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, LSNCDNM, INWON_STAT, "
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
    
    sStr = sStr & "    FROM (SELECT ACID, LSNCD,"
    sStr = sStr & "                 GET_LSNNM(ACID, LSNCD) AS LSNNM,"
    sStr = sStr & "                 GET_LSNCDNM(ACID, LSNCD) AS LSNCDNM,"
    
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
    
    sStr = sStr & "                 SUM(SEL_X2) AS SEL_X2,"

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
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN 1 "
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                        0 "
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
    sStr = sStr & "                                       '외국어'"                                 '< 변경
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       ''"
    sStr = sStr & "                               END SEL_N3,"
    sStr = sStr & "                               CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                                       ''"                                       '< 변경
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
    sStr = sStr & "      ORDER BY LSNCDNM "
    
    
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
            
            sprTotGwamok.MaxRows = .RecordCount + 1
                sprTotGwamok.Row = 1:           sprTotGwamok.RowHeight(sprTotGwamok.Row) = nRowHeight
                
            Call Exec_sprTotGwamok_Formula          '< 합계처리
                
            For nRec = 2 To .RecordCount + 1 Step 1
                
                sprTotGwamok.Row = nRec:            sprTotGwamok.RowHeight(sprTotGwamok.Row) = nRowHeight
                
                sprTotGwamok.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprTotGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprTotGwamok.Col = sprTotGwamok.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprTotGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprTotGwamok.Col = sprTotGwamok.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprTotGwamok, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprTotGwamok.SetCellBorder sprTotGwamok.Col, 1, sprTotGwamok.Col, sprTotGwamok.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> 총인원
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                sprTotGwamok.SetCellBorder sprTotGwamok.Col, 1, sprTotGwamok.Col, sprTotGwamok.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
                
                '<< 인문자연 공통 : 8 과목 >>
                For nCol = 1 To 8 Step 1
                    sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '사탐은 9~11
                        For nCol = 9 To 11 Step 1
                            sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            sprTotGwamok.CellType = CellTypeNumber
                            sprTotGwamok.TypeVAlign = TypeVAlignCenter
                            sprTotGwamok.TypeNumberDecPlaces = 0
                            sprTotGwamok.TypeNumberMin = -9999
                            sprTotGwamok.TypeNumberMax = 9999
                            
                            sprTotGwamok.TypeNumberShowSep = False
                            If nTmp > 0 Then sprTotGwamok.value = nTmp
                            
                        Next nCol
                        
                    Case "02"
                        '과탐은 COLUMN만 이동
                        For nCol = 9 To 11 Step 1
                            sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                            sprTotGwamok.CellType = CellTypeNumber
                            sprTotGwamok.TypeVAlign = TypeVAlignCenter
                            sprTotGwamok.TypeNumberDecPlaces = 0
                            sprTotGwamok.TypeNumberMin = -9999
                            sprTotGwamok.TypeNumberMax = 9999
                            
                            sprTotGwamok.TypeNumberShowSep = False
                            If nTmp > 0 Then sprTotGwamok.value = nTmp
                            
                        Next nCol
                End Select
                
                sprTotGwamok.SetCellBorder sprTotGwamok.Col, 1, sprTotGwamok.Col, sprTotGwamok.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 제2선택
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_X2")) = False Then
                        nTmp = CDbl(.Fields("SEL_X2"))
                    End If
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                sprTotGwamok.SetCellBorder sprTotGwamok.Col, 1, sprTotGwamok.Col, sprTotGwamok.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 언
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N1")) = False Then
                        nTmp = CDbl(.Fields("SEL_N1"))
                    End If
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                '> 수
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N2")) = False Then
                        nTmp = CDbl(.Fields("SEL_N2"))
                    End If
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                '> 사
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N3")) = False Then
                        nTmp = CDbl(.Fields("SEL_N3"))
                    End If
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                '> 탐
                sprTotGwamok.Col = sprTotGwamok.Col + 1:    nTmp = 0
                    If IsNull(.Fields("SEL_N4")) = False Then
                        nTmp = CDbl(.Fields("SEL_N4"))
                    End If
                    
                    sprTotGwamok.CellType = CellTypeNumber
                    sprTotGwamok.TypeVAlign = TypeVAlignCenter
                    sprTotGwamok.TypeNumberDecPlaces = 0
                    sprTotGwamok.TypeNumberMin = -9999
                    sprTotGwamok.TypeNumberMax = 9999
                    
                    sprTotGwamok.TypeNumberShowSep = False
                    If nTmp > 0 Then sprTotGwamok.value = nTmp
                    
                
                sprTotGwamok.SetCellBorder sprTotGwamok.Col, 1, sprTotGwamok.Col, sprTotGwamok.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprTotGwamok.Col = sprTotGwamok.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprTotGwamok)
                    sprTotGwamok.value = 0
                
                .MoveNext       '<< 다음항목
                
            Next nRec
            
            sprTotGwamok.Row = 1:       sprTotGwamok.Row2 = sprTotGwamok.MaxRows
            sprTotGwamok.Col = 1:       sprTotGwamok.Col2 = sprTotGwamok.MaxCols
            sprTotGwamok.BlockMode = True
                sprTotGwamok.BackColor = basModule.WhiteColor
                sprTotGwamok.BackColorStyle = BackColorStyleUnderGrid
            sprTotGwamok.BlockMode = False

            sprTotGwamok.ColsFrozen = 4
            
            sprTotGwamok.Row = 1:       sprTotGwamok.Row2 = 1
            sprTotGwamok.Col = 1:       sprTotGwamok.Col2 = sprTotGwamok.MaxCols
            sprTotGwamok.BlockMode = True
                sprTotGwamok.BackColor = &H80C0FF
                sprTotGwamok.BackColorStyle = BackColorStyleUnderGrid
            sprTotGwamok.BlockMode = False
            
            sprTotGwamok.SetCellBorder 1, 1, sprTotGwamok.MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
            
        '>> spread lock
            sprTotGwamok.Row = 1:       sprTotGwamok.Row2 = sprTotGwamok.MaxRows
            sprTotGwamok.Col = 1:       sprTotGwamok.Col2 = sprTotGwamok.MaxCols
            sprTotGwamok.BlockMode = True
                sprTotGwamok.Lock = True
                sprTotGwamok.Protect = True
            sprTotGwamok.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "반별 수강신청내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "반별 수강신청내역 조회"
    
End Sub






Private Sub sprSTD_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then
        
        With sprSTD
            ' Sort by ZIP Code in descending order
            .SortKey(1) = Col
            .SortKeyOrder(1) = SortKeyOrderAscending
            .Sort -1, -1, -1, -1, SortByRow

        End With
        
    End If
End Sub

'>> 선택 ## multi 선택
Private Sub sprTotGwamok_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    
    If Row < 2 Then Exit Sub

    With sprTotGwamok
        If .MaxRows < 1 Then Exit Sub

        sprTotGwamok.Enabled = False
        
            If .Tag = "0" Then
                .Row = CLng(.Tag):      .Row2 = .Row
                .Col = 1:               .Col2 = .MaxCols
                .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = basModule.WhiteColor
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
                .BackColor = basModule.SelectColor1
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
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                Else
                    .value = 1
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = basModule.SelectColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    .Tag = Trim(CStr(Row))
                End If
            
            End If
            
        sprTotGwamok.Enabled = True

    End With
End Sub

Private Sub sprTotGwamok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nS      As Long
    Dim nE      As Long
    
    Dim nRow    As Long
    
    With sprTotGwamok
    
        If .MaxRows = 0 Then Exit Sub
        
        Select Case Shift
'            Case 0
'                Call sprTotGwamok_Click(1, .ActiveRow)
                
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
                            .BackColor = basModule.SelectColor1
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








'## 반 선택하기
Private Sub cmdGetLsn_Click()
    
    Dim nRowTotGwamok       As Long
    Dim nRowBanChk          As Long
    
    Dim nCol                As Long
    
    Dim sTmp                As String
    Dim nTmp                As Long
    
    Dim bAddChk             As Boolean
    
    Dim sLsnCD              As String
    Dim sLsnCDNM            As String
    Dim sC_LsnCD            As String
    Dim sC_LsnCDNM          As String
    
    Dim sAdd_LsnCD          As String       '< 학생 반 조회시 사용
    
    cmdGetLsn.Enabled = False
    
    
    For nRowTotGwamok = sprTotGwamok.MaxRows To 2 Step -1
        bAddChk = False
        
        sprTotGwamok.Row = nRowTotGwamok
        sprTotGwamok.Col = sprTotGwamok.MaxCols
        
        If sprTotGwamok.value = 1 Then
            sprTotGwamok.Col = 1:       sLsnCD = Trim(sprTotGwamok.Text)
            sprTotGwamok.Col = 3:       sLsnCDNM = Trim(sprTotGwamok.Text)
            
            '< 기존 있는지 검토함. >
            For nRowBanChk = 1 To sprBanChk.MaxRows Step 1
                sprBanChk.Row = nRowBanChk
                sprBanChk.Col = 1
                
                If StrComp(sLsnCD, sprBanChk.Text, vbTextCompare) = 0 Then
                    lblStatus.Caption = "이미 선택된 반입니다."
                    GoTo NextRow
                End If
            Next nRowBanChk
            
            '< 기존 선택된 내용이 아님. > => sprBanChk에 ADD
            '  단, 어디에 넣을지 선택
            For nRowBanChk = sprBanChk.MaxRows To 2 Step -1
                
                sprBanChk.Row = nRowBanChk
                sprBanChk.Col = 1:      sC_LsnCD = Trim(sprBanChk.Text)
                sprBanChk.Col = 3:      sC_LsnCDNM = Trim(sprBanChk.Text)
                
                If StrComp(sLsnCD, sC_LsnCD, vbTextCompare) > 0 And _
                   StrComp(sLsnCDNM, sC_LsnCDNM, vbTextCompare) > 0 Then
                   
                   sprBanChk.MaxRows = sprBanChk.MaxRows + 1
                    sprBanChk.InsertRows nRowBanChk + 1, 1
                        sprBanChk.Row = nRowBanChk + 1:     sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight
                   
                   bAddChk = True
                   Exit For
                End If
            Next nRowBanChk
            
            If bAddChk = False Then
                sprBanChk.MaxRows = sprBanChk.MaxRows + 1
                sprBanChk.InsertRows 2, 1
                    sprBanChk.Row = 2:      sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight
                
            End If
            
            bAddChk = False
            
            '< Data ADD >
            sprTotGwamok.Row = nRowTotGwamok
                
                For nCol = 1 To 3 Step 1
                    sprTotGwamok.Col = nCol
                        sTmp = Trim(sprTotGwamok.Text)
                        
                        sprBanChk.Col = nCol
                            Call basFunction.Set_SprType_Text(sprBanChk, "center", "left", basFunction.LenKor(sTmp), sTmp)
                Next nCol
                                
                For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                    sprTotGwamok.Col = nCol
                        nTmp = 0:       If Trim(sprTotGwamok.Text) <> "" Then nTmp = sprTotGwamok.value
                    
                        sprBanChk.Col = nCol
                            sprBanChk.CellType = CellTypeNumber
                            sprBanChk.TypeVAlign = TypeVAlignCenter
                            sprBanChk.TypeNumberDecPlaces = 0
                            sprBanChk.TypeNumberMin = -9999
                            sprBanChk.TypeNumberMax = 9999
                            
                            sprBanChk.TypeNumberShowSep = False
                                                        
                            If nTmp > 0 Then sprBanChk.value = nTmp
                Next nCol
            
            bAddChk = True
            
        End If
NextRow:
        
    Next nRowTotGwamok
    
    '< 합계처리 >
    With sprBanChk
        
        For nCol = 4 To (4 + 11 + 1 + 4) Step 1
            sprBanChk.Row = 1
            sprBanChk.Col = nCol
                sprBanChk.Text = ""
        Next nCol
        
        sAdd_LsnCD = ""
        
        For nRowBanChk = 2 To .MaxRows Step 1
            .Row = nRowBanChk
            .Col = 1
                If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
            
            For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                
                nTmp = 0
                
                .Row = nRowBanChk
                .Col = nCol
                If Trim(.Text) <> "" Then
                    If .BackColor = basModule.WhiteColor Then
                        nTmp = .value
                    End If
                End If
                    
                If nTmp > 0 Then
                    .Row = 1
                    .Col = nCol
                    
                        sprBanChk.CellType = CellTypeNumber
                        sprBanChk.TypeVAlign = TypeVAlignCenter
                        sprBanChk.TypeNumberDecPlaces = 0
                        sprBanChk.TypeNumberMin = -9999
                        sprBanChk.TypeNumberMax = 9999
                        
                        sprBanChk.TypeNumberShowSep = False
                        
                    If Trim(.Text) = "" Then
                        .value = nTmp
                    Else
                        .value = .value + nTmp
                    End If
                End If
                   
            Next nCol
        Next nRowBanChk
        
        .SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        .SetCellBorder 4, 1, 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11, 1, 4 + 11, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11 + 1, 1, 4 + 11 + 1, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11 + 1 + 4, 1, 4 + 11 + 1 + 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        
        .AddCellSpan 1, 1, 3, 1
        .Row = 1
        .Col = 1
            .Text = "합  계"
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .ForeColor = basModule.SectionColor1
        
        .Row = 1:   .Row2 = .Row
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = &HFFC0C0
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = 1:   .Row2 = .MaxCols
        .Col = 4:   .Col2 = 4
        .BlockMode = True
            .BackColor = &HFFC0C0
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        sprBanChk.SetCellBorder 1, 1, sprBanChk.MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
        
        '<< lock >>
        For nRowBanChk = 1 To sprBanChk.MaxRows Step 1
            sprBanChk.Row = nRowBanChk
            sprBanChk.Col = 1
            
            If Trim(sprBanChk.Text) < "90000" Then
                sprBanChk.Row2 = sprBanChk.Row
                sprBanChk.Col = 1:      sprBanChk.Col2 = sprBanChk.MaxCols
                
                sprBanChk.BlockMode = True
                    sprBanChk.Lock = True
                    sprBanChk.Protect = True
                sprBanChk.BlockMode = False
                
            End If
        Next nRowBanChk
        
        
        If sAdd_LsnCD > " " Then
            Call Find_STD_Data(sAdd_LsnCD)         '< 선택 반의 학생조회
            MsgBox "조회 완료하였습니다.", vbInformation + vbOKOnly, "작업 반 선택"
        End If
        
    End With
    
    cmdGetLsn.Enabled = True
    
End Sub


Private Sub Exec_sprBanChk_Formula()
    Dim nCol        As Long
    
    With sprBanChk

     '>> 열 합계 -------------------------------------------------------
            For nCol = 4 To (.MaxCols - 1) Step 1
                .Row = 1
                .Col = nCol
                
                .CellType = CellTypeNumber
                .TypeVAlign = TypeVAlignCenter
                .TypeNumberDecPlaces = 0
                .TypeNumberMin = -9999
                .TypeNumberMax = 9999
                
                .TypeNumberShowSep = False
            Next nCol
            
            For nCol = 4 To .MaxCols Step 1
                .Row = 1
                .Col = nCol
                .FormulaSync = True
                .Formula = "SUM(#2:#" & Trim(CStr(.MaxRows)) & ")"
                
            Next nCol
            
    End With
    
End Sub

'## 학생의 신청내용 조회
Private Sub Find_STD_Data(ByVal aAdd_LsnCD As String)
    
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
    
    sprSTD.MaxRows = 0
    
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
    sStr = sStr & "                     '외국어'"                               '< 변경
    sStr = sStr & "                 ELSE"
    sStr = sStr & "                     ''"
    sStr = sStr & "                 END SEL_N3,"
    sStr = sStr & "                 CASE WHEN INSTR(SEL5,'94|') > 0 THEN"
    sStr = sStr & "                     ''"                                     '< 변경
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
    
    sStr = sStr & "             AND SEL_CLASS IN (" & aAdd_LsnCD & ") "
    
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
    sStr = sStr & "   ORDER BY SEL_CLASS, GAEYUL_CD, EXMID, STDNM"
    
    
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
                sprSTD.MaxRows = sprSTD.MaxRows + 1
                sprSTD.Row = sprSTD.MaxRows ':      SPRSTD.RowHeight(SPRSTD.Row) = nRowHeight

                sprSTD.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid


                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMTYPE")) = False Then sTmp = Trim(.Fields("EXMTYPE"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("EXMTYPE_NM")) = False Then sTmp = Trim(.Fields("EXMTYPE_NM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)

                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS")) = False Then sTmp = Trim(.Fields("SEL_CLASS"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ":     If IsNull(.Fields("SEL_CLASS_NM")) = False Then sTmp = Trim(.Fields("SEL_CLASS_NM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                For ni = 1 To 4 Step 1
                    sFieldNM = ""

                    sFieldNM = "GWA_BAN" & Trim(CStr(ni))
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni
                
                sprSTD.Col = sprSTD.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprSTD)
                    sprSTD.value = 0

                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                For ni = 1 To 11 Step 1
                    sFieldNM = ""

                    sFieldNM = "SEL" & Trim(CStr(ni))
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ":     If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                Next ni

                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("SEL_X2")) = False Then sTmp = Trim(.Fields("SEL_X2"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)

                
                .MoveNext       '<< 다음항목
                
            Next nRec
        End If
        
        With sprSTD
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




'## sort
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
    
    With sprSTD
        For ni = 1 To 16 Step 1
            For nj = 1 To 16 Step 1
            
                If fpSort(nj).value = ni Then
                    nC = nC + 1
                    
                    .SortKey(nC) = nj + 13 - 1
                    .SortKeyOrder(nC) = SortKeyOrderAscending
                End If
            Next nj
        Next ni
        
        .Sort -1, -1, -1, -1, SortByRow
        
    End With

End Sub






'## 반 학생선택
Private Sub sprBanChk_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nColor      As Long
    
    Dim nTotTmp     As Long
    Dim nTmp        As Long
    
    Dim nCol        As Long
    Dim sAdd_LsnCD  As String
    Dim nRowBanChk  As Long
    Dim sLsnCD      As String
    
    lblStatus.Caption = ""
    
    If Row < 2 Then
        lblStatus.Caption = "과목을 선택하세요."
        Exit Sub
    End If
    
    If Col <= 4 Then
        lblStatus.Caption = "과목을 선택하세요."
        Exit Sub
    End If
    
    nColor = basModule.WhiteColor
    
    With sprBanChk
        .Row = Row
        .Col = Col
        
        If .BackColor <> basModule.WhiteColor Then
            lblStatus.Caption = "이미 등록된 내용이 있습니다."
            Exit Sub
        Else
            If optLsn(0).value = True Then
                nColor = optLsn(0).BackColor
            ElseIf optLsn(1).value = True Then
                nColor = optLsn(1).BackColor
            ElseIf optLsn(2).value = True Then
                nColor = optLsn(2).BackColor
            ElseIf optLsn(3).value = True Then
                nColor = optLsn(3).BackColor
            End If
            
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = nColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            
            '- 반 인원 다시 계산 -----------------------------------------------------------------------
            For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                sprBanChk.Row = 1
                sprBanChk.Col = nCol
                    sprBanChk.Text = ""
            Next nCol
            
            sAdd_LsnCD = ""
    
            For nRowBanChk = 2 To .MaxRows Step 1
                .Row = nRowBanChk
                .Col = 1
                    If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                    sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
                
                    sLsnCD = Trim(.Text)
                
                For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                    
                    nTmp = 0
                    
                    .Row = nRowBanChk
                    .Col = nCol
                    If Trim(.Text) <> "" Then
                        If .BackColor = basModule.WhiteColor Then
                            nTmp = .value
                        Else
                            If nCol >= 5 And nCol <= .MaxCols Then
                                If sLsnCD >= "90000" Then
                                    nTmp = -1 * .value
                                End If
                            End If
                        End If
                    End If
                        
                    If nTmp > 0 Then
                        .Row = 1
                        .Col = nCol
                        
                            sprBanChk.CellType = CellTypeNumber
                            sprBanChk.TypeVAlign = TypeVAlignCenter
                            sprBanChk.TypeNumberDecPlaces = 0
                            sprBanChk.TypeNumberMin = -9999
                            sprBanChk.TypeNumberMax = 9999
                            
                            sprBanChk.TypeNumberShowSep = False
                            
                        If Trim(.Text) = "" Then
                            .value = nTmp
                        Else
                            .value = .value + nTmp
                        End If
                    End If
                       
                Next nCol
            Next nRowBanChk
            '----------------------------------------------------------------------------
                
        End If
    End With
    
End Sub


'## 반 등록내역 삭제
Private Sub sprBanChk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nColor          As Long
    
    Dim nTotTmp         As Long
    Dim nTmp            As Long
    Dim nRowBanChk      As Long
    Dim nCol            As Long
    
    Dim sAdd_LsnCD      As String
    Dim sLsnCD          As String
    
    lblStatus.Caption = ""
    
    If sprBanChk.ActiveRow < 2 Then
        lblStatus.Caption = "과목을 선택하세요."
        Exit Sub
    End If
    
    If sprBanChk.ActiveCol <= 4 Then
        lblStatus.Caption = "과목을 선택하세요."
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        With sprBanChk
            .Row = .ActiveRow
            .Col = .ActiveCol
            
            If .BackColor = basModule.WhiteColor Then
                lblStatus.Caption = "삭제항목이 없습니다."
                Exit Sub
            Else
                nColor = basModule.WhiteColor
                
                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = nColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = 1
                If Trim(.Text) >= "90000" Then
                    .Col = .ActiveCol
                        .Text = ""
                End If
                
                
                '- 반 인원 다시 계산 -----------------------------------------------------------------------
                For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                    sprBanChk.Row = 1
                    sprBanChk.Col = nCol
                        sprBanChk.Text = ""
                Next nCol
                
                sAdd_LsnCD = ""
        
                For nRowBanChk = 2 To .MaxRows Step 1
                    .Row = nRowBanChk
                    .Col = 1
                        If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                        sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
                    
                        sLsnCD = Trim(.Text)
                    
                    For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                        
                        nTmp = 0
                        
                        .Row = nRowBanChk
                        .Col = nCol
                        If Trim(.Text) <> "" Then
                            If .BackColor = basModule.WhiteColor Then
                                nTmp = .value
                            Else
                                If nCol >= 5 And nCol <= .MaxCols Then
                                    If sLsnCD >= "90000" Then
                                        nTmp = -1 * .value
                                    End If
                                End If
                            End If
                        End If
                            
                        If nTmp <> 0 Then
                            .Row = 1
                            .Col = nCol
                            
                                sprBanChk.CellType = CellTypeNumber
                                sprBanChk.TypeVAlign = TypeVAlignCenter
                                sprBanChk.TypeNumberDecPlaces = 0
                                sprBanChk.TypeNumberMin = -9999
                                sprBanChk.TypeNumberMax = 9999
                                
                                sprBanChk.TypeNumberShowSep = False
                                
                            If Trim(.Text) = "" Then
                                .value = nTmp
                            Else
                                .value = .value + nTmp
                            End If
                        End If
                           
                    Next nCol
                Next nRowBanChk
                '----------------------------------------------------------------------------
                
            End If
        End With
        
    End If
    
End Sub


'## 1. 반 삭제처리
'## 2. key 입력
Private Sub sprBanChk_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTmp            As String
    Dim nTmp            As Long
    
    Dim nCol            As Long
    Dim nRowBanChk      As Long
    Dim sAdd_LsnCD      As String
    
    Dim nColor          As Long
    Dim sLsnCD          As String
    
    lblStatus.Caption = ""
    
    Select Case KeyCode
        Case vbKeyNumpad0 To vbKeyNumpad9, vbKey1 To vbKey9, vbKeyBack
            With sprBanChk
                If .ActiveCol <= 4 Then
                    lblStatus.Caption = "과목에서 인원수를 등록하세요."
                    Exit Sub
                End If
                
                If .ActiveRow < 2 Then
                    lblStatus.Caption = "과목에서 인원수를 등록하세요."
                    Exit Sub
                End If
                
                .Row = .ActiveRow
                .Col = .ActiveCol
                
                If optLsn(0).value = True Then
                    nColor = optLsn(0).BackColor
                ElseIf optLsn(1).value = True Then
                    nColor = optLsn(1).BackColor
                ElseIf optLsn(2).value = True Then
                    nColor = optLsn(2).BackColor
                ElseIf optLsn(3).value = True Then
                    nColor = optLsn(3).BackColor
                End If
                
                If Trim(.Text) = "" Then
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                ElseIf Trim(.Text) = "0" Then
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = basModule.WhiteColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                Else
                    .Row2 = .Row
                    .Col2 = .Col
                    .BlockMode = True
                        .BackColor = nColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                End If
                
                
                '- 반 인원 다시 계산 -----------------------------------------------------------------------
                For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                    .Row = 1
                    .Col = nCol
                        .Text = ""
                Next nCol
                
                sAdd_LsnCD = ""
        
                For nRowBanChk = 2 To .MaxRows Step 1
                    .Row = nRowBanChk
                    .Col = 1
                        If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                        sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
                    
                        sLsnCD = Trim(.Text)
                    
                    For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                        
                        nTmp = 0
                        
                        .Row = nRowBanChk
                        .Col = nCol
                        If Trim(.Text) <> "" Then
                            If .BackColor = basModule.WhiteColor Then
                                nTmp = .value
                            Else
                                If nCol >= 5 And nCol <= .MaxCols Then
                                    If sLsnCD >= "90000" Then
                                        nTmp = -1 * .value
                                    End If
                                End If
                            End If
                        End If
                            
                        If nTmp <> 0 Then
                            .Row = 1
                            .Col = nCol
                            
                                sprBanChk.CellType = CellTypeNumber
                                sprBanChk.TypeVAlign = TypeVAlignCenter
                                sprBanChk.TypeNumberDecPlaces = 0
                                sprBanChk.TypeNumberMin = -9999
                                sprBanChk.TypeNumberMax = 9999
                                
                                sprBanChk.TypeNumberShowSep = False
                                
                            If Trim(.Text) = "" Then
                                .value = nTmp
                            Else
                                .value = .value + nTmp
                            End If
                        End If
                           
                    Next nCol
                Next nRowBanChk
        
                .SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
                .SetCellBorder 4, 1, 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11, 1, 4 + 11, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11 + 1, 1, 4 + 11 + 1, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11 + 1 + 4, 1, 4 + 11 + 1 + 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                .AddCellSpan 1, 1, 3, 1
                .Row = 1
                .Col = 1
                    .Text = "합  계"
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    .ForeColor = basModule.SectionColor1
                
                .Row = 1:   .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = &HFFC0C0
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = 1:   .Row2 = .MaxCols
                .Col = 4:   .Col2 = 4
                .BlockMode = True
                    .BackColor = &HFFC0C0
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                sprBanChk.SetCellBorder 1, 1, sprBanChk.MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
                
                '----------------------------------------------------------------------------
                
            End With
            
        Case vbKeyDelete
            With sprBanChk
                .Row = .ActiveRow
                .Col = 2
                    sTmp = "【 " & Trim(.Text) & " 】의 반을 삭제하시겠습니까?"
                
                If MsgBox(sTmp, vbQuestion + vbYesNo, "반 삭제") = vbNo Then
                    lblStatus.Caption = "반 삭제취소하였습니다."
                    Exit Sub
                End If
                
                '<< 반 삭제 >>
                .DeleteRows .ActiveRow, 1
                .MaxRows = .MaxRows - 1
                
                '- 반 인원 다시 계산 -----------------------------------------------------------------------
                For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                    .Row = 1
                    .Col = nCol
                        .Text = ""
                Next nCol
                
                sAdd_LsnCD = ""
        
                For nRowBanChk = 2 To .MaxRows Step 1
                    .Row = nRowBanChk
                    .Col = 1
                        If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                        sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
                    
                        sLsnCD = Trim(.Text)
                    
                    For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                        
                        nTmp = 0
                        
                        .Row = nRowBanChk
                        .Col = nCol
                        If Trim(.Text) <> "" Then
                            If .BackColor = basModule.WhiteColor Then
                                nTmp = .value
                            Else
                                If nCol >= 5 And nCol <= .MaxCols Then
                                    If sLsnCD >= "90000" Then
                                        nTmp = -1 * .value
                                    End If
                                End If
                            End If
                        End If
                            
                        If nTmp > 0 Then
                            .Row = 1
                            .Col = nCol
                            
                                sprBanChk.CellType = CellTypeNumber
                                sprBanChk.TypeVAlign = TypeVAlignCenter
                                sprBanChk.TypeNumberDecPlaces = 0
                                sprBanChk.TypeNumberMin = -9999
                                sprBanChk.TypeNumberMax = 9999
                                
                                sprBanChk.TypeNumberShowSep = False
                                
                            If Trim(.Text) = "" Then
                                .value = nTmp
                            Else
                                .value = .value + nTmp
                            End If
                        End If
                           
                    Next nCol
                Next nRowBanChk
        
                .SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
                .SetCellBorder 4, 1, 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11, 1, 4 + 11, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11 + 1, 1, 4 + 11 + 1, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                .SetCellBorder 4 + 11 + 1 + 4, 1, 4 + 11 + 1 + 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                .AddCellSpan 1, 1, 3, 1
                .Row = 1
                .Col = 1
                    .Text = "합  계"
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    .ForeColor = basModule.SectionColor1
                
                .Row = 1:   .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = &HFFC0C0
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Row = 1:   .Row2 = .MaxCols
                .Col = 4:   .Col2 = 4
                .BlockMode = True
                    .BackColor = &HFFC0C0
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                sprBanChk.SetCellBorder 1, 1, sprBanChk.MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
                    
                If sAdd_LsnCD > " " Then
                    Call Find_STD_Data(sAdd_LsnCD)         '< 선택 반의 학생조회
                    MsgBox "조회 완료하였습니다.", vbInformation + vbOKOnly, "작업 반 선택"
                End If
                '----------------------------------------------------------------------------
                
            End With
            
    End Select
End Sub






'## 반별 선택과목 등록하기
Private Sub cmdinPut_Click()
    Dim nRow            As Long
    Dim nCnt_Lsn        As String
    Dim sTmp            As String
    
    cmdinPut.Enabled = False
    
        With sprBanChk
            nCnt_Lsn = 0
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                If Trim(.Text) < "90000" Then
                    nCnt_Lsn = nCnt_Lsn + 1
                End If
            Next nRow
            
            If nCnt_Lsn = 0 Then
                MsgBox "등록된 반이 없거나 이동반만 있습니다.", vbExclamation + vbOKOnly, "선택과목 등록"
                cmdinPut.Enabled = True
                Exit Sub
            End If
        End With
        
        sTmp = ""
        sTmp = "【 " & Trim(Left(cboKaeyol.Text, 30)) & " 】계열 "
        sTmp = sTmp & "【 " & Trim(Left(cboLsnType.Text, 30))
        sTmp = sTmp & " 】타입으로 현 선택과목 내역을 등록하시겠습니까?"
        If MsgBox(sTmp, vbQuestion + vbYesNo, "선택과목 등록") = vbNo Then
            cmdinPut.Enabled = True
            Exit Sub
        End If
        
        Call Save_inPutData
        
    cmdinPut.Enabled = True
End Sub

Private Sub Save_inPutData()
    
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nTot        As Long
    Dim nExeTot     As Long
    Dim nExe        As Long
    Dim nLength     As Long
    
    Dim nRow        As Long
    Dim nCol        As Integer
    Dim ni          As Integer
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
'>> 등록방법 : 기존의 등록된 type 에 해당하는 내역을 모두 삭제 후 처리함.
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    '<< TYPE 에 해당하는 내역을 모두 삭제 >>
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDLSN05TB "
    sStr = sStr & "  WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "    AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
'    '>> ACID
'        sTmp = Trim(basModule.SchCD)
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
    
    
    '<< 행의 값을 모두 저장 >>
    With sprBanChk
        nTot = 0
        nExeTot = 0
        nExe = 0
        
        For nRow = 2 To .MaxRows Step 1
            nTot = nTot + 1
            
            .Row = nRow
            
            sStr = ""
            sStr = sStr & " INSERT INTO SDLSN05TB ( "
            sStr = sStr & "        ACID       , KAEYOL     , LSNTYPE    , LSNCD      ,"
            sStr = sStr & "        LSN_STD_SUM,"
            
            sStr = sStr & "        TAMGU1     , TAMGU2     , TAMGU3     , TAMGU4     , TAMGU5     ,"
            sStr = sStr & "        TAMGU6     , TAMGU7     , TAMGU8     , TAMGU9     , TAMGU10    ,"
            sStr = sStr & "        TAMGU11    ,"
            sStr = sStr & "        NONSUL1    , NONSUL2    , NONSUL3    , NONSUL4    , "
            sStr = sStr & "        J2SEL      ,"
            
            sStr = sStr & "        TAMGU_CL1  , TAMGU_CL2  , TAMGU_CL3  , TAMGU_CL4  , TAMGU_CL5  ,"
            sStr = sStr & "        TAMGU_CL6  , TAMGU_CL7  , TAMGU_CL8  , TAMGU_CL9  , TAMGU_CL10 ,"
            sStr = sStr & "        TAMGU_CL11 ,"
            
            sStr = sStr & "        NONSUL1_CL , NONSUL2_CL , NONSUL3_CL , NONSUL4_CL , "
            sStr = sStr & "        J2SEL_CL   "
            sStr = sStr & " ) "
            sStr = sStr & " VALUES ( "
            sStr = sStr & "       '" & Trim(basModule.SchCD) & "', "
            sStr = sStr & "       '" & Trim(Right(cboKaeyol.Text, 30)) & "', "
            sStr = sStr & "       '" & Trim(Right(cboLsnType.Text, 30)) & "', "
            .Col = 1:       sTmp = Trim(.Text)
                sStr = sStr & "   '" & sTmp & "', "                     '< LSNCD
            .Col = 4:       nTmp = 0:       If IsNumeric(.Text) = True Then nTmp = CLng(.Text)
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< LSN_STD_SUM : 반 전체인원
            
        '/* 탐구 */
            ni = 5
            For nCol = 0 To 10 Step 1       '< 탐구영역 11과목
                .Col = ni + nCol
                    nTmp = 0:       If IsNumeric(.Text) = True Then nTmp = CLng(.Text)
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< 탐구영역
            Next nCol
            
        '/* 논술 */
            ni = 5 + 11 + 1
            For nCol = 0 To 3 Step 1        '< 논술 4과목
                .Col = ni + nCol
                    nTmp = 0:       If IsNumeric(.Text) = True Then nTmp = CLng(.Text)
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< 논술영역
            Next nCol
        '/* 제2영역 */
            .Col = 5 + 11:          nTmp = 0:       If IsNumeric(.Text) = True Then nTmp = CLng(.Text)
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< 제2영역
            
        '/* 탐구 backcolor */
            ni = 5
            For nCol = 0 To 10 Step 1       '< 탐구영역 11과목
                .Col = ni + nCol
                    nTmp = .BackColor
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< 탐구영역
            Next nCol
        '/* 논술 backcolor */
            ni = 5 + 11 + 1
            For nCol = 0 To 3 Step 1        '< 논술 4과목
                .Col = ni + nCol
                    nTmp = .BackColor
                sStr = sStr & "    " & Trim(CStr(nTmp)) & ", "         '< 논술영역
            Next nCol
            
        '/* 제2영역 backcolor */
            .Col = 5 + 11
                nTmp = .BackColor
                sStr = sStr & "    " & Trim(CStr(nTmp))                 '< 제2영역
            
            sStr = sStr & " )"
            
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
    
            nExe = 0
            DBCmd.Execute nExe, , -1
    
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
    
            If nExe = 1 Then
                nExeTot = nExeTot + 1
            End If
            
        Next nRow
    End With
    
    '>> 처리수가 동일해야 함.
    If nTot = nExeTot Then
        basDataBase.DBConn.CommitTrans
        MsgBox "선택과목 등록하였습니다.", vbInformation + vbOKOnly, "선택과목 등록"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "등록 중 에러가 발생하였습니다.", vbCritical + vbOKOnly, "선택과목 등록"
    End If
    
    ' NO ERROR
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "선택과목 등록 중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & " " & Err.Description, vbCritical + vbOKOnly, "선택과목 등록"
    
    On Error GoTo 0
End Sub






'## 기 등록내용 가져오기
Private Sub cmdSearchSaveData_Click()
    Dim sTmp        As String
    
    cmdSearchSaveData.Enabled = False
        
        
        sTmp = ""
        sTmp = "【 " & Trim(Left(cboKaeyol.Text, 30)) & " 】계열 "
        sTmp = sTmp & "【 " & Trim(Left(cboLsnTypeCP.Text, 30))
        sTmp = sTmp & " 】타입을 조회하시겠습니까?"
        If MsgBox(sTmp, vbQuestion + vbYesNo, "선택과목 조회") = vbNo Then
            cmdSearchSaveData.Enabled = True
            Exit Sub
        End If
        
        Call SearchSaveData
        
        sprBanChk.SetFocus
        sprBanChk.SetActiveCell 1, 2
        
    cmdSearchSaveData.Enabled = True
    
End Sub

Private Sub SearchSaveData()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Long
    Dim siTem       As String
    Dim sAdd_LsnCD  As String
    Dim nRowBanChk  As Long
    
    sprBanChk.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = sStr & "    SELECT"
    sStr = sStr & "        A.ACID     , A.LSNTYPE  , "
    sStr = sStr & "        A.LSNCD    , B.LSNNM    , B.LSNCDNM  , B.KAEYOL   , "
    sStr = sStr & "        LSN_STD_SUM,"
    
    sStr = sStr & "        TAMGU1     , TAMGU2     , TAMGU3     , TAMGU4     , TAMGU5     ,"
    sStr = sStr & "        TAMGU6     , TAMGU7     , TAMGU8     , TAMGU9     , TAMGU10    ,"
    sStr = sStr & "        TAMGU11    ,"
    sStr = sStr & "        J2SEL      ,"
    sStr = sStr & "        NONSUL1    , NONSUL2    , NONSUL3    , NONSUL4    ,"
    
    sStr = sStr & "        TAMGU_CL1  , TAMGU_CL2  , TAMGU_CL3  , TAMGU_CL4  , TAMGU_CL5  ,"
    sStr = sStr & "        TAMGU_CL6  , TAMGU_CL7  , TAMGU_CL8  , TAMGU_CL9  , TAMGU_CL10 ,"
    sStr = sStr & "        TAMGU_CL11 ,"
    sStr = sStr & "        J2SEL_CL   ,"
    sStr = sStr & "        NONSUL1_CL , NONSUL2_CL , NONSUL3_CL , NONSUL4_CL"
    sStr = sStr & "      FROM SDLSN05TB A, "
    sStr = sStr & "           ("
    sStr = sStr & "            SELECT *"
    sStr = sStr & "              From SDLSN01TB"
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            Union All"
    sStr = sStr & "            SELECT *"
    sStr = sStr & "              From SDLSN02TB"
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            ) B "
    sStr = sStr & "     WHERE A.ACID    = B.ACID "
    sStr = sStr & "       AND A.LSNCD   = B.LSNCD "
    sStr = sStr & "       AND A.ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "       AND A.KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "       AND A.LSNTYPE = '" & Trim(Right(cboLsnTypeCP.Text, 30)) & "'"
    
    
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
            
            sprBanChk.MaxRows = .RecordCount + 1
                sprBanChk.Row = 1:           sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight
                
            Call Exec_sprBanChk_Formula          '< 합계처리
                
            For nRec = 2 To .RecordCount + 1 Step 1
                
                sprBanChk.Row = nRec:            sprBanChk.RowHeight(sprBanChk.Row) = nRowHeight
                
                sprBanChk.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprBanChk.Col = sprBanChk.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprBanChk.Col = sprBanChk.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        Call basFunction.Set_SprType_Text(sprBanChk, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> 총인원
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("LSN_STD_SUM")) = False Then
                        nTmp = CDbl(.Fields("LSN_STD_SUM"))
                    End If
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
                
                '<< 인문자연 공통 : 8 과목 >>
                For nCol = 1 To 8 Step 1
                    sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    siTem = "TAMGU" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    nTmp = basModule.WhiteColor
                    siTem = "TAMGU_CL" & Trim(CStr(nCol))
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '사탐은 9~11
                        For nCol = 9 To 11 Step 1
                            sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                            siTem = "TAMGU" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            sprBanChk.CellType = CellTypeNumber
                            sprBanChk.TypeVAlign = TypeVAlignCenter
                            sprBanChk.TypeNumberDecPlaces = 0
                            sprBanChk.TypeNumberMin = -9999
                            sprBanChk.TypeNumberMax = 9999
                            
                            sprBanChk.TypeNumberShowSep = False
                            If nTmp > 0 Then sprBanChk.value = nTmp
                                                        
                                                        
                            nTmp = basModule.WhiteColor
                            siTem = "TAMGU_CL" & Trim(CStr(nCol))
                                If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                                
                            sprBanChk.Row2 = sprBanChk.Row
                            sprBanChk.Col2 = sprBanChk.Col
                            sprBanChk.BlockMode = True
                                sprBanChk.BackColor = nTmp
                                sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                            sprBanChk.BlockMode = False
                        Next nCol
                        
                    Case "02"
                        '과탐은 COLUMN만 이동
                        For nCol = 9 To 11 Step 1
                            sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                            sprBanChk.CellType = CellTypeNumber
                            sprBanChk.TypeVAlign = TypeVAlignCenter
                            sprBanChk.TypeNumberDecPlaces = 0
                            sprBanChk.TypeNumberMin = -9999
                            sprBanChk.TypeNumberMax = 9999
                            
                            sprBanChk.TypeNumberShowSep = False
                            If nTmp > 0 Then sprBanChk.value = nTmp
                            
                            
                            nTmp = basModule.WhiteColor
                            siTem = "TAMGU_CL" & Trim(CStr(nCol))
                                If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                                
                            sprBanChk.Row2 = sprBanChk.Row
                            sprBanChk.Col2 = sprBanChk.Col
                            sprBanChk.BlockMode = True
                                sprBanChk.BackColor = nTmp
                                sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                            sprBanChk.BlockMode = False
                            
                        Next nCol
                End Select
                
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 제2선택
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("J2SEL")) = False Then
                        nTmp = CDbl(.Fields("J2SEL"))
                    End If
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    
                    nTmp = basModule.WhiteColor
                    siTem = "J2SEL_CL"
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                    
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                '> 언
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("NONSUL1")) = False Then
                        nTmp = CDbl(.Fields("NONSUL1"))
                    End If
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    nTmp = basModule.WhiteColor
                    siTem = "NONSUL1_CL"
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                    
                '> 수
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("NONSUL2")) = False Then
                        nTmp = CDbl(.Fields("NONSUL2"))
                    End If
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    nTmp = basModule.WhiteColor
                    siTem = "NONSUL2_CL"
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                    
                '> 사
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("NONSUL3")) = False Then
                        nTmp = CDbl(.Fields("NONSUL3"))
                    End If
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    nTmp = basModule.WhiteColor
                    siTem = "NONSUL3_CL"
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                    
                '> 탐
                sprBanChk.Col = sprBanChk.Col + 1:    nTmp = 0
                    If IsNull(.Fields("NONSUL4")) = False Then
                        nTmp = CDbl(.Fields("NONSUL4"))
                    End If
                    
                    sprBanChk.CellType = CellTypeNumber
                    sprBanChk.TypeVAlign = TypeVAlignCenter
                    sprBanChk.TypeNumberDecPlaces = 0
                    sprBanChk.TypeNumberMin = -9999
                    sprBanChk.TypeNumberMax = 9999
                    
                    sprBanChk.TypeNumberShowSep = False
                    If nTmp > 0 Then sprBanChk.value = nTmp
                    
                    nTmp = basModule.WhiteColor
                    siTem = "NONSUL4_CL"
                        If IsNumeric(.Fields(siTem)) = True Then nTmp = CDbl(.Fields(siTem))
                        
                    sprBanChk.Row2 = sprBanChk.Row
                    sprBanChk.Col2 = sprBanChk.Col
                    sprBanChk.BlockMode = True
                        sprBanChk.BackColor = nTmp
                        sprBanChk.BackColorStyle = BackColorStyleUnderGrid
                    sprBanChk.BlockMode = False
                
                sprBanChk.SetCellBorder sprBanChk.Col, 1, sprBanChk.Col, sprBanChk.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                .MoveNext       '<< 다음항목
                
            Next nRec
        End If
    End With

    If sprBanChk.MaxRows = 0 Then
        Call cboKaeyol_Click
        lblStatus.Caption = "기 등록된 내용이 없습니다.[전체 반 선택과목 신청조회]을 조회 후 [반 선택하기] 하세요."
        Exit Sub
    End If
    

    '< 합계처리 >
    With sprBanChk
        
        For nCol = 4 To (4 + 11 + 1 + 4) Step 1
            sprBanChk.Row = 1
            sprBanChk.Col = nCol
                sprBanChk.Text = ""
        Next nCol
        
        sAdd_LsnCD = ""
        
        For nRowBanChk = 2 To .MaxRows Step 1
            .Row = nRowBanChk
            .Col = 1
                If sAdd_LsnCD > " " Then sAdd_LsnCD = sAdd_LsnCD & ","
                sAdd_LsnCD = sAdd_LsnCD & "'" & Trim(.Text) & "'"
            
            For nCol = 4 To (4 + 11 + 1 + 4) Step 1
                
                nTmp = 0
                
                .Row = nRowBanChk
                .Col = nCol
                If Trim(.Text) <> "" Then
                    If .BackColor = basModule.WhiteColor Then
                        nTmp = .value
                    End If
                End If
                    
                If nTmp > 0 Then
                    .Row = 1
                    .Col = nCol
                    
                        sprBanChk.CellType = CellTypeNumber
                        sprBanChk.TypeVAlign = TypeVAlignCenter
                        sprBanChk.TypeNumberDecPlaces = 0
                        sprBanChk.TypeNumberMin = -9999
                        sprBanChk.TypeNumberMax = 9999
                        
                        sprBanChk.TypeNumberShowSep = False
                        
                    If Trim(.Text) = "" Then
                        .value = nTmp
                    Else
                        .value = .value + nTmp
                    End If
                End If
                   
            Next nCol
        Next nRowBanChk
        
        .SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
        .SetCellBorder 4, 1, 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11, 1, 4 + 11, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11 + 1, 1, 4 + 11 + 1, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        .SetCellBorder 4 + 11 + 1 + 4, 1, 4 + 11 + 1 + 4, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
        
        .AddCellSpan 1, 1, 3, 1
        .Row = 1
        .Col = 1
            .Text = "합  계"
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .ForeColor = basModule.SectionColor1
        
        .Row = 1:   .Row2 = .Row
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = &HFFC0C0
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = 1:   .Row2 = .MaxCols
        .Col = 4:   .Col2 = 4
        .BlockMode = True
            .BackColor = &HFFC0C0
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        sprBanChk.SetCellBorder 1, 1, sprBanChk.MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
        
        '<< lock >>
        For nRowBanChk = 1 To sprBanChk.MaxRows Step 1
            sprBanChk.Row = nRowBanChk
            sprBanChk.Col = 1
            
            If Trim(sprBanChk.Text) < "90000" Then
                sprBanChk.Row2 = sprBanChk.Row
                sprBanChk.Col = 1:      sprBanChk.Col2 = sprBanChk.MaxCols
                
                sprBanChk.BlockMode = True
                    sprBanChk.Lock = True
                    sprBanChk.Protect = True
                sprBanChk.BlockMode = False
                
            End If
        Next nRowBanChk
        
        
        If sAdd_LsnCD > " " Then
            Call Find_STD_Data(sAdd_LsnCD)         '< 선택 반의 학생조회
            MsgBox "조회 완료하였습니다.", vbInformation + vbOKOnly, "기 등록내역 가져오기"
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
    MsgBox "선택과목 내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "기 등록내역 가져오기"
    
    
End Sub




Private Sub cmdBanToGwamok_Click()
    Load TMR028
    
    '## type 보내주기
    Call TMR028.init_Data(Trim(Right(cboKaeyol.Text, 30)), Trim(Right(cboLsnTypeCP.Text, 30)))
    
    TMR028.Show
    TMR028.ZOrder 0

End Sub














