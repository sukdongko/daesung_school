VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form STD020 
   Caption         =   "���л��� >> �հݻ� ���"
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
      BorderStyle     =   0  '����
      Caption         =   "Frame4"
      Height          =   8325
      Left            =   30
      TabIndex        =   42
      Top             =   1140
      Width           =   15045
      Begin VB.Frame Frame5 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
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
               Name            =   "����"
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
            Caption         =   "�հ����"
            Height          =   500
            Left            =   480
            TabIndex        =   62
            Top             =   7680
            Width           =   1905
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "���"
            Height          =   225
            Left            =   14010
            TabIndex        =   30
            Top             =   600
            Width           =   675
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "�հݻ� ����ϱ� (&S)"
            Height          =   500
            Left            =   12540
            TabIndex        =   32
            Top             =   7680
            Width           =   2000
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame6"
            Height          =   525
            Left            =   30
            TabIndex        =   50
            Top             =   30
            Width           =   9675
            Begin VB.CommandButton cmdSort 
               BackColor       =   &H00C0C0FF&
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               BackStyle       =   0  '����
               Caption         =   "���ŵ��"
               Height          =   210
               Left            =   3150
               TabIndex        =   67
               Top             =   30
               Width           =   825
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�ܱ���"
               Height          =   210
               Index           =   3
               Left            =   6510
               TabIndex        =   61
               Top             =   15
               Width           =   615
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Index           =   2
               Left            =   5700
               TabIndex        =   60
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "���"
               Height          =   210
               Index           =   1
               Left            =   4830
               TabIndex        =   59
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Index           =   0
               Left            =   8970
               TabIndex        =   44
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label11 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�迭"
               Height          =   210
               Left            =   8220
               TabIndex        =   45
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label10 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�հ�"
               Height          =   210
               Left            =   7410
               TabIndex        =   46
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label9 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "��.��"
               Height          =   210
               Left            =   4050
               TabIndex        =   47
               Top             =   15
               Width           =   465
            End
            Begin VB.Label Label8 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   2070
               TabIndex        =   48
               Top             =   15
               Width           =   405
            End
            Begin VB.Label Label7 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�����ȣ"
               Height          =   210
               Left            =   1050
               TabIndex        =   49
               Top             =   15
               Width           =   795
            End
            Begin VB.Label Label5 
               BackStyle       =   0  '����
               Caption         =   "> ����"
               BeginProperty Font 
                  Name            =   "����"
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
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   29
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   2
            Left            =   11880
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   28
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   1
            Left            =   10800
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   27
            Top             =   240
            Width           =   1035
         End
         Begin VB.ComboBox cboPass 
            Height          =   300
            Index           =   0
            Left            =   9720
            Style           =   2  '��Ӵٿ� ���
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
            BackStyle       =   0  '����
            Caption         =   ">> �հ��п� ------------------------------"
            BeginProperty Font 
               Name            =   "����"
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
      BorderStyle     =   0  '����
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   30
      TabIndex        =   33
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   1005
         Left            =   30
         TabIndex        =   34
         Top             =   30
         Width           =   14985
         Begin VB.ComboBox cboSel2_Sch 
            Height          =   300
            Left            =   13530
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   172
            Width           =   1395
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   4140
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   2
            Top             =   165
            Width           =   1005
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ȸ�ϱ� (&F)"
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
               Caption         =   "�հ�ó����"
               Height          =   285
               Left            =   90
               TabIndex        =   8
               Top             =   150
               Width           =   1365
            End
            Begin VB.OptionButton optPassY 
               BackColor       =   &H00D2EAF5&
               Caption         =   "�հݻ���"
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
            Style           =   2  '��Ӵٿ� ���
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
            Style           =   2  '��Ӵٿ� ���
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
               Name            =   "����"
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
               Name            =   "����"
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
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               BackStyle       =   0  '����
               Caption         =   "�ܱ���            �̻�/"
               Height          =   210
               Left            =   3420
               TabIndex        =   58
               Top             =   180
               Width           =   1755
            End
            Begin VB.Label Label16 
               BackStyle       =   0  '����
               Caption         =   "����            �̻�/"
               Height          =   210
               Left            =   1770
               TabIndex        =   57
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label15 
               BackStyle       =   0  '����
               Caption         =   "���            �̻�/"
               Height          =   210
               Left            =   150
               TabIndex        =   56
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label6 
               BackStyle       =   0  '����
               Caption         =   "�հ�             ����"
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
               Name            =   "����"
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
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "2���� �п�"
            Height          =   210
            Left            =   12570
            TabIndex        =   65
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '����
            Caption         =   "��ȸ�ο�"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   13350
            TabIndex        =   64
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   3690
            TabIndex        =   53
            Top             =   210
            Width           =   405
         End
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����п�"
            Height          =   210
            Left            =   1560
            TabIndex        =   40
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ             ����             ����"
            Height          =   210
            Left            =   5460
            TabIndex        =   39
            Top             =   210
            Width           =   3075
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�������"
            Height          =   210
            Left            =   10020
            TabIndex        =   38
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л���"
            Height          =   210
            Left            =   8100
            TabIndex        =   37
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��  ��"
            Height          =   210
            Left            =   2370
            TabIndex        =   36
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  '����
            Caption         =   ">> ��ȸ�׸�"
            BeginProperty Font 
               Name            =   "����"
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
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : STD020
'   �� ��  �� �� : �հݻ� ���
'
'   ��   ��   �� : 2007/08/24
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private sini_Path       As String    '>> �뼺�п�
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
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            
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
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "������" & Space(30) & "0"
            .AddItem "������" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        '��2����
        With cboSel2_Sch
            .Clear
            .AddItem "����" & Space(30) & "X"
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            
            .ListIndex = 0
        End With
    
'>> �迭
        With cboKaeyol
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
        '<< �迭 >> : 2008.01.09
            If Trim(basModule.SchCD) = "N" Then             '< �뷮��
                .AddItem "��ü" & Space(30) & "03"
                .AddItem "����(��)" & Space(30) & "04"
                .AddItem "�ι�����" & Space(30) & "05"
                .AddItem "�ڿ�����" & Space(30) & "06"
                
                .AddItem "�ι�-��" & Space(30) & "07"
                .AddItem "�ڿ�-��" & Space(30) & "08"
                '.AddItem "�����ι�-��" & Space(30) & "09"
                '.AddItem "�����ڿ�-��" & Space(30) & "10"
                
                .AddItem "��)�ι�" & Space(30) & "11"
                .AddItem "��)�ڿ�" & Space(30) & "12"
                .AddItem "��)��ü" & Space(30) & "13"
                .AddItem "��)����(��)" & Space(30) & "14"
                .AddItem "��)�ι�����" & Space(30) & "15"
                .AddItem "��)�ڿ�����" & Space(30) & "16"
            End If
        '<< �迭 >> : 2008.01.10
            If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then           '< ���� 2008.03.24
                .AddItem "�ָ�����" & Space(30) & "04"
                .AddItem "�ָ��Ǵ�" & Space(30) & "05"
                
                .AddItem "�߰�����" & Space(30) & "06"
                .AddItem "�߰��Ǵ�" & Space(30) & "07"
            
                .AddItem "�������ι�" & Space(30) & "11"
                .AddItem "�������ڿ�" & Space(30) & "12"
                
                .AddItem "�������ι�16" & Space(30) & "16"
                .AddItem "�������ڿ�17" & Space(30) & "17"
                
            End If
        '<< �迭 >> : 2008.02.15
            If Trim(basModule.SchCD) = "S" Then             '< ����
''                .AddItem "��ü��" & Space(30) & "03"
''
''                .AddItem "�ι�����" & Space(30) & "05"
''                .AddItem "�ڿ�����" & Space(30) & "06"
''
                .AddItem "�ż��ι�" & Space(30) & "11"
                .AddItem "�ż��ڿ�" & Space(30) & "12"
                
                .AddItem "�ι������̾�" & Space(30) & "18"
                .AddItem "�ڿ������̾�" & Space(30) & "19"
                
               .AddItem "�����Ư���ι�" & Space(30) & "21"
               .AddItem "�����Ư���ڿ�" & Space(30) & "22"
               .AddItem "�߰�������ι�" & Space(30) & "23"
               .AddItem "�߰�������ڿ�" & Space(30) & "24"
             
                
            End If
        '<< �迭 >> : 2008.02.15
            If Trim(basModule.SchCD) = "P" Then             '< ����
                .AddItem "Ư���ι�" & Space(30) & "03"
                .AddItem "Ư���ڿ�" & Space(30) & "04"
            End If
            
            If Trim(basModule.SchCD) = "J" Then             '< ����
                .AddItem "�ż��ι�" & Space(30) & "11"
                .AddItem "�ż��ڿ�" & Space(30) & "12"
                
                .AddItem "�ι������̾�" & Space(30) & "18"
                .AddItem "�ڿ������̾�" & Space(30) & "19"
            End If
            
            
        '<< �迭 >> : 2009.01.09
            If Trim(basModule.SchCD) = "B" Then             '< �λ�
                .AddItem "���м����ι�" & Space(30) & "05"
                .AddItem "���м����ڿ�" & Space(30) & "06"
                
                .AddItem "��.����ι�" & Space(30) & "07"
                .AddItem "��.����ڿ�" & Space(30) & "08"
                
                .AddItem "��ȭ�ι�" & Space(30) & "09"
                .AddItem "��ȭ�ڿ�" & Space(30) & "10"
            End If
            
            .ListIndex = 0
        End With
        
            
        With cboPass(0)
            .Clear
            .AddItem "����" & Space(30) & "X"
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(1)
            .Clear
            .AddItem "����" & Space(30) & "X"
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(2)
            .Clear
            .AddItem "����" & Space(30) & "X"
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            .ListIndex = 0
        End With
        With cboPass(3)
            .Clear
            .AddItem "����" & Space(30) & "X"
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
            .AddItem "�ָ����Ǵ�" & Space(30) & "W"
            .AddItem "�߰����Ǵ�" & Space(30) & "Q"
            
            .AddItem "����" & Space(30) & "J"
            .AddItem "�λ�" & Space(30) & "B"
            .AddItem "�������(��õ)" & Space(30) & "E"
            .ListIndex = 0
        End With
            
        sprPass.Tag = "0"           '<< ������ ������ ���ؼ�
        
        sini_Path = App.Path & "\DAESUNG.INI"
        
        '>> ���α׷� INI ����
        If Dir(sini_Path) = "" Then                                     '<< ������ ������ ����
            sSort = insert_Order_ini_File("0,6/1,5/2,2/3,3/4,1/5,4/")
        End If
        
        sGbn = "STD020"
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SORT", "", sData, 255, sini_Path)         '>> SORT ����
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
    
    
'    fpSort(0).Value = 6         '<< �л�
'    fpSort(1).Value = 5         '<< ����
'    fpSort(2).Value = 2         '<< ��.��
'    fpSort(3).Value = 3         '<< �հ�
'    fpSort(4).Value = 1         '<< �迭
'    fpSort(5).Value = 4         '<< ����

'    fpSort(6).Value = 3         '<< ���
'    fpSort(7).Value = 3         '<< ����
'    fpSort(8).Value = 3         '<< �ܱ���
'    fpSort(9).Value = 1         '<< ���ŵ��

End Sub

Private Function insert_Order_ini_File(ByVal aSort As String) As String
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sReturn     As String
    
    sGbn = "STD020"
         sReturn = basModule.WritePrivateProfileString(sGbn, "SORT", aSort, sini_Path)                 '<< SORT ����
    
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
                        Case 0                      '<< �����ȣ
                            .SortKey(nC) = 3
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "0," & CInt(Trim(fpSort(0).value)) & "/"
                        Case 1                      '<< ����
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "1," & CInt(Trim(fpSort(1).value)) & "/"
                        Case 2                      '<< ��.������
                            .SortKey(nC) = 6
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "2," & CInt(Trim(fpSort(2).value)) & "/"
                        Case 3                      '<< �հ�
                            .SortKey(nC) = 11
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "3," & CInt(Trim(fpSort(3).value)) & "/"
                        Case 4                      '<< �迭
                            .SortKey(nC) = 13
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "4," & CInt(Trim(fpSort(4).value)) & "/"
                        Case 5                      '<< ����
                            .SortKey(nC) = 14
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "5," & CInt(Trim(fpSort(5).value)) & "/"
                            
                            
                        Case 6                      '<< ���
                            .SortKey(nC) = 8
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "6," & CInt(Trim(fpSort(6).value)) & "/"
                        Case 7                      '<< ����
                            .SortKey(nC) = 9
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "7," & CInt(Trim(fpSort(7).value)) & "/"
                        Case 8                      '<< �ܱ���
                            .SortKey(nC) = 10
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "8," & CInt(Trim(fpSort(8).value)) & "/"
                        Case 9                      '<< ���ŵ��
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










'>> ��ȸ������ �л��˻�
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
    sStr = sStr & "         SEL2_SCH, "                                     '< 2008.01.11 : ���� M -> ��2����
    
    sStr = sStr & "         KAEYOL_CD, KAEYOL_NM, "
    sStr = sStr & "         PASS1, PASS2, PASS3, PASS4 "
    sStr = sStr & "    FROM ("
            sStr = sStr & "  SELECT SCHNO, ACID, EXMID, STDNM, SUBSTR(Birth_ymd,1,4)||'-'||SUBSTR(Birth_ymd,5,2)  ||'-'||SUBSTR(Birth_ymd,7,2) AS Birth_ymd_F, Birth_ymd ,"
            sStr = sStr & "         EXMTYPE, DECODE(EXMTYPE,'0','������','1','������') AS EXMTYPE_NM,"
            sStr = sStr & "         K_NUM, E_NUM, M_NUM, N_NUM,"
            sStr = sStr & "         NVL( NVL(K_NUM,0)+NVL(E_NUM,0)+NVL(M_NUM,0), 0) AS  TOT_NUM,"
            sStr = sStr & "         SEL1, SEL3, "
            
            
'            sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
'            sStr = sStr & "             '��Ž'"
'            sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
'            sStr = sStr & "             '��Ž'"
'            sStr = sStr & "         END END GAEYUL,"
            sStr = sStr & "         SEL2_SCH, "                             '< 2008.01.11 : ���� M -> ��2����
            
            sStr = sStr & "         KAEYOL AS KAEYOL_CD,"
            
            '<< �迭 >> : 2008.01.09
            If Trim(basModule.SchCD) = "N" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                sStr = sStr & "                   '03','��ü',"
                sStr = sStr & "                   '04','����(��)',"
                sStr = sStr & "                   '05','�ι�����',"
                sStr = sStr & "                   '06','�ڿ�����',"
                
                sStr = sStr & "                   '07','�ż��ι�',"
                sStr = sStr & "                   '08','�ż��ڿ�',"
                sStr = sStr & "                   '09','�ż������ι�',"
                sStr = sStr & "                   '10','�ż������ڿ�',"
                
                sStr = sStr & "                   '11','��)�ι�',"
                sStr = sStr & "                   '12','��)�ڿ�',"
                sStr = sStr & "                   '13','��)��ü',"
                sStr = sStr & "                   '14','��)����(��)',"
                sStr = sStr & "                   '15','��)�ι�����',"
                sStr = sStr & "                   '16','��)�ڿ�����'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< �迭 >> : 2008.01.10
            ElseIf Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                
                sStr = sStr & "                   '04','�ָ�����',"
                sStr = sStr & "                   '05','�ָ��Ǵ�',"
                sStr = sStr & "                   '06','�߰�����',"
                sStr = sStr & "                   '07','�߰��Ǵ�',"
                
                sStr = sStr & "                   '11','�������ι�',"
                sStr = sStr & "                   '12','�������ڿ�',"
                
                sStr = sStr & "                   '16','�������ι�16',"
                sStr = sStr & "                   '17','�������ڿ�17'"
                
                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< �迭 >> : 2008.02.15
            ElseIf Trim(basModule.SchCD) = "S" Then
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                sStr = sStr & "                   '03','��ü��',"
                
                sStr = sStr & "                   '05','�����ι�',"
                sStr = sStr & "                   '06','�����ڿ�',"
                
                sStr = sStr & "                   '11','�ż��ι�',"
                sStr = sStr & "                   '12','�ż��ڿ�',"
                
                sStr = sStr & "                   '18','�ι������̾�',"
                sStr = sStr & "                   '19','�ڿ������̾�',"
                
                sStr = sStr & "                   '21','�����Ư���ι�',"
                sStr = sStr & "                   '22','�����Ư���ڿ�',"
                
                sStr = sStr & "                   '23','�߰�������ι�',"
                sStr = sStr & "                   '24','�߰�������ڿ�'"
                

                sStr = sStr & "            ) AS KAEYOL_NM,"
            '<< �迭 >> : 2008.02.15
            ElseIf Trim(basModule.SchCD) = "P" Then         '< ����
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                sStr = sStr & "                   '03','Ư���ι�',"
                sStr = sStr & "                   '04','Ư���ڿ�'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            ElseIf Trim(basModule.SchCD) = "J" Then         '< ����
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                sStr = sStr & "                   '11','�ż��ι�',"
                sStr = sStr & "                   '12','�ż��ڿ�',"
                
                sStr = sStr & "                   '18','�ι������̾�',"
                sStr = sStr & "                   '19','�ڿ������̾�'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            ElseIf Trim(basModule.SchCD) = "B" Then         '< �λ� : 2009.01.09
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�',"
                sStr = sStr & "                   '05','Ư���ι�',"
                sStr = sStr & "                   '06','Ư���ڿ�',"
                sStr = sStr & "                   '07','������ι�',"
                sStr = sStr & "                   '08','������ڿ�',"
                sStr = sStr & "                   '09','��ȭ�ι�',"
                sStr = sStr & "                   '10','��ȭ�ڿ�'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
                
            Else
                sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
                sStr = sStr & "                   '02','�ڿ�'"
                sStr = sStr & "            ) AS KAEYOL_NM,"
            End If
            
            sStr = sStr & "         PASS1, PASS2, PASS3, PASS4 "
            sStr = sStr & "    From CLSTD01TB"
            sStr = sStr & "   WHERE ACID  = '" & Trim(Right(cboHakwon.Text, 30)) & "'"
            sStr = sStr & "     AND EXMID > ' ' "           '> ������ �л�
            sStr = sStr & "     AND CL_CLOSE IS NULL "      '> �ϷῩ�� : ����Ǹ� YYMM���� ��.
            
            sStr = sStr & "     AND BIGO2 IS NULL"          '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
    
    
    Select Case basModule.SchCD
        Case "K"
            sStr = sStr & "     AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            
        Case Else
            If optPassN.value = True Then
                'If Trim(basModule.SchCD) = "N" Then
                    'sStr = sStr & "     AND BIGO1 = '7' "                                      '> �ϷῩ�� : ����Ǹ� YYMM���� ��.
                    'sStr = sStr & "     AND REGDATE <= TO_DATE('20080315','YYYYMMDD')"         '< ���Ǳ�� �� �� !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    
                    'sStr = sStr & "     AND BIGO1 = (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB WHERE ACID = '" & Trim(basModule.SchCD) & "') "      '> �ϷῩ�� : ����Ǹ� YYMM���� ��. : 2008.02.28
                    sStr = sStr & "     AND BIGO1 = (SELECT TO_CHAR(MAX(TO_NUMBER(BIGO1))) FROM CLSTD01TB ) "      '> �ϷῩ�� : ����Ǹ� YYMM���� ��. : 2008.02.28
                'End If
            End If
            
    End Select
       
    sStr = sStr & "          )"
    sStr = sStr & "   WHERE SCHNO > ' ' "
    
'>> �հ�ó���� �л� & �հݵ� �л�
    If optPassN.value = True Then
        sStr = sStr & " AND (PASS1 IS NULL AND PASS2 IS NULL AND PASS3 IS NULL AND PASS4 IS NULL)"
    ElseIf optPassY.value = True Then
        sStr = sStr & " AND (PASS1 > ' ' OR PASS2 > ' ' OR PASS3 > ' ' OR PASS4 > ' ') "
    End If
    
'>> ����
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "ALL"
            ' NO ACTION
        Case "0"
            sStr = sStr & " AND EXMTYPE = '0' "     '<< ������
        Case "1"
            sStr = sStr & " AND EXMTYPE = '1' "     '<< ������
    End Select
    
'>> ��2����
    Select Case Trim(Right(cboSel2_Sch.Text, 30))
        Case "X"
            ' no action
        Case Else
            sStr = sStr & " AND SEL2_SCH = '" & Trim(Right(cboSel2_Sch.Text, 30)) & "'"
    End Select
    
'>> �迭
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
    
'>> �����ȣ
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = "" Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) = "" Then
        ' no action
    End If
       
'>> �л���
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & " AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
    End If
'>> �ֹι�ȣ
    If Trim(fpBirth_ymd.UnFmtText) > " " Then
        sStr = sStr & " AND Birth_ymd LIKE '" & Trim(fpBirth_ymd.UnFmtText) & "%'"
    End If
    
'>> �հ�
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
            '>> ���
                If fpKor.value > 0 Then
                    sStr = sStr & " AND K_NUM <= " & Trim(CStr(fpKor.value))
                End If
            '>> ����
                If fpMat.value > 0 Then
                    sStr = sStr & " AND M_NUM <= " & Trim(CStr(fpMat.value))
                End If
            '>> �ܱ���
                If fpEng.value > 0 Then
                    sStr = sStr & " AND E_NUM <= " & Trim(CStr(fpEng.value))
                End If
        Case "1"
            '>> ���
                If fpKor.value > 0 Then
                    sStr = sStr & " AND K_NUM >= " & Trim(CStr(fpKor.value))
                End If
            '>> ����
                If fpMat.value > 0 Then
                    sStr = sStr & " AND M_NUM >= " & Trim(CStr(fpMat.value))
                End If
            '>> �ܱ���
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
    

    
'    '>> �п�
'        sTmp = Trim(Right(cboHakwon.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �����ȣ
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
'    '>> �հ�
'        If fpTot.Value > 0 Then
'            nTmp = CLng(fpTot.Value)
'            Set DBParam = DBCmd.CreateParameter("TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'        End If
'    '>> �л���
'        If Trim(txtStdNM.Text) > " " Then
'            sTmp = "%" & Trim(txtStdNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'    '>> �ֹι�ȣ
'        If Trim(fpBirth_ymd.UnFmtText) > " " Then
'            sTmp = "%" & Trim(fpBirth_ymd.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
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
                        
                sprPass.Col = sprPass.Col + 1       ' �����ȣ
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
                    nTmp = 0:   If IsNull(.Fields("N_NUM")) = False Then nTmp = CDbl(Trim(.Fields("N_NUM"))) '���ŵ��
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
                    sTmp = " ": If IsNull(.Fields("SEL2_SCH")) = False Then sTmp = Trim(.Fields("SEL2_SCH"))        '< 2008.01.11 : ���� M -> ��2����
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "�뷮��"
                        Case "K"
                            sTmp = "����"
                        Case "S"
                            sTmp = "����"
                        Case "P"
                            sTmp = "���� M"
                        Case "M"
                            sTmp = "���� M"
                            
                        Case "W"
                            sTmp = "�ָ����Ǵ�"
                        Case "Q"
                            sTmp = "�߰����Ǵ�"
                            
                        Case "J"
                            sTmp = "����"
                        Case "B"
                            sTmp = "�λ�"
                       Case "E"
                            sTmp = "�������(��õ)"
                    End Select
                    Call basFunction.Set_SprType_Text(sprPass, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                
                If IsNull(.Fields("PASS1")) = True Then
                    sprPass.Col = sprPass.Col + 1
                Else
                    sprPass.Col = sprPass.Col + 1
                    sTmp = Trim(.Fields("PASS1"))
                    Select Case UCase(Trim(sTmp))
                        Case "N"
                            sTmp = "�뷮��" & Space(30) & "N"
                        Case "K"
                            sTmp = "����" & Space(30) & "K"
                        Case "S"
                            sTmp = "����" & Space(30) & "S"
                        Case "P"
                            sTmp = "���� M" & Space(30) & "P"
                        Case "M"
                            sTmp = "���� M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "�ָ����Ǵ�" & Space(30) & "W"
                        Case "Q"
                            sTmp = "�߰����Ǵ�" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "����" & Space(30) & "J"
                        Case "B"
                            sTmp = "�λ�" & Space(30) & "B"
                        Case "E"
                            sTmp = "�������(��õ)" & Space(30) & "E"
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
                            sTmp = "�뷮��" & Space(30) & "N"
                        Case "K"
                            sTmp = "����" & Space(30) & "K"
                        Case "S"
                            sTmp = "����" & Space(30) & "S"
                        Case "P"
                            sTmp = "���� M" & Space(30) & "P"
                        Case "M"
                            sTmp = "���� M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "�ָ����Ǵ�" & Space(30) & "W"
                        Case "Q"
                            sTmp = "�߰����Ǵ�" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "����" & Space(30) & "J"
                        Case "B"
                            sTmp = "�λ�" & Space(30) & "B"
                         Case "E"
                            sTmp = "�������(��õ)" & Space(30) & "E"
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
                            sTmp = "�뷮��" & Space(30) & "N"
                        Case "K"
                            sTmp = "����" & Space(30) & "K"
                        Case "S"
                            sTmp = "����" & Space(30) & "S"
                        Case "P"
                            sTmp = "���� M" & Space(30) & "P"
                        Case "M"
                            sTmp = "���� M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "�ָ����Ǵ�" & Space(30) & "W"
                        Case "Q"
                            sTmp = "�߰����Ǵ�" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "����" & Space(30) & "J"
                        Case "B"
                            sTmp = "�λ�" & Space(30) & "B"
                        Case "E"
                            sTmp = "�������(��õ)" & Space(30) & "E"
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
                            sTmp = "�뷮��" & Space(30) & "N"
                        Case "K"
                            sTmp = "����" & Space(30) & "K"
                        Case "S"
                            sTmp = "����" & Space(30) & "S"
                        Case "P"
                            sTmp = "���� M" & Space(30) & "P"
                        Case "M"
                            sTmp = "���� M" & Space(30) & "M"
                            
                        Case "W"
                            sTmp = "�ָ����Ǵ�" & Space(30) & "W"
                        Case "Q"
                            sTmp = "�߰����Ǵ�" & Space(30) & "Q"
                            
                        Case "J"
                            sTmp = "����" & Space(30) & "J"
                        Case "B"
                            sTmp = "�λ�" & Space(30) & "B"
                        Case "E"
                            sTmp = "�������(��õ)" & Space(30) & "E"
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
    
    MsgBox "�л� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л� �հ� �� Ȯ��"
    
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
    MsgBox "�հ�ó�� �� Ȯ�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� �հ� �� Ȯ��"
End Sub
















'>> ���� ## multi ����
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
                    If .Tag > "0" Then              '<< 1. �����ϰ� 2. shift�� ���� ��Ƽ�� ������ ���
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

'>> ��ü����
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









































'======================================================== �հ��л� ����ϱ� ============================================================

'>> �հ��п� �ֱ�
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
            MsgBox "�հ��л��� �հ��п��� �����Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "�հ��� ó��"
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



'>> �հݻ� ����ϱ�
Private Sub cmdSave_Click()
    Dim bRet        As Boolean
    
    Dim ni      As Long
    Dim nRec    As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    With sprPass
        If .MaxRows = 0 Then Exit Sub
        
'>> üũ����
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
            MsgBox "�հ��л��� �����Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "�հ��� ó��"
            Exit Sub
        End If
    End With
    
    On Error GoTo ErrStmt
    
    cmdSave.Enabled = False
        bRet = Save_STD_Data
        
    cmdSave.Enabled = True
    
    If bRet = True Then
        MsgBox "�հ��� ��� �Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "�հ��� ó��"
    Else
        MsgBox "�հ��� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�հ��� ó��"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�հ��� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�հ��� ó��"
    On Error GoTo 0
    
End Sub





'>> �л��ڵ尡 �����ϹǷ�
'>> �հ��� ó���� �л��ڵ�θ� ������Ʈ �մϴ�.
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
    
    Dim nRec        As Long         '<< ó���ؾ� �� ��
    Dim nTot        As Long         '<< ó���� ��
    
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
            
            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
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
            
            '>> 1�� �հ�
                sprPass.Col = 16
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 2�� �հ�
                sprPass.Col = 17
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 3�� �հ�
                sprPass.Col = 18
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> 4�� �հ�
                sprPass.Col = 19
                sTmp = Trim(Right(sprPass.Text, 30))
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                    
            '>> �л��ڵ�
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





'>> ���ó��
Private Sub cmdCancel_Click()
    Dim nRow        As Long
    Dim nCnt        As Long
    
    Dim nChkRow     As Long
    
    Dim sAcID       As String
    Dim sSchNO      As String
    Dim sExmID      As String
    
    
    If optPassY.value = False Then
        MsgBox "�հݻ��� ��ȸ�ÿ� �����մϴ�.", vbExclamation + vbOKOnly, "���ó��"
        Exit Sub
    End If
    
    If sprPass.MaxRows = 0 Then
        MsgBox "�հݻ��� ��ȸ�� ó���Ͻʽÿ�.", vbExclamation + vbOKOnly, "���ó��"
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
    
    MsgBox "�Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "���ó��"
    
'    If nCnt > 1 Or nCnt = 0 Then
'        MsgBox "�հݻ� 1�� ���ó�������մϴ�.", vbExclamation + vbOKOnly, "���ó��"
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
    
    '1 �п�
    '2 �л�
    '3 �����ȣ
    
    cmdCancel.Enabled = False
        bRet = Cancel_StdOut(aAcID, aSchNO, aExmID)
        
    cmdCancel.Enabled = True
    
    If bRet = True Then
        'MsgBox "�л� �հ���� �Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л� ����ϱ�"
    Else
        MsgBox "�л� �հ���ҽ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� ����ϱ�"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л� �հ���ҽ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� ����ϱ�"
    On Error GoTo 0
    
End Sub

'>> �հ�����ϱ�
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
