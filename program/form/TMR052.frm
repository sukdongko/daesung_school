VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR052 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  '���� ����
   Caption         =   "�ð�ǥ ����� >> ��ü�ð�ǥ ���� >> �ð�ǥ ����"
   ClientHeight    =   13245
   ClientLeft      =   7170
   ClientTop       =   765
   ClientWidth     =   14790
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13245
   ScaleWidth      =   14790
   Begin VB.Frame fraTmrChg 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  '����
      Height          =   13125
      Left            =   90
      TabIndex        =   21
      Top             =   90
      Width           =   14565
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   13065
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   14505
         Begin VB.CommandButton cmd_P_TmrChg 
            Caption         =   "���泻�� ����ϱ� (&S)"
            Height          =   555
            Left            =   11280
            TabIndex        =   16
            Top             =   2910
            Width           =   2385
         End
         Begin FPSpread.vaSpread sprTmr 
            Height          =   9135
            Left            =   210
            TabIndex        =   41
            Top             =   3810
            Width           =   14085
            _Version        =   393216
            _ExtentX        =   24844
            _ExtentY        =   16113
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
            SpreadDesigner  =   "TMR052.frx":0000
         End
         Begin VB.Frame fraFrom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "������ ����   [����Ű�� ó��]"
            Height          =   2745
            Left            =   210
            TabIndex        =   29
            Top             =   150
            Width           =   4905
            Begin VB.CommandButton cmdTmr 
               Caption         =   "�ش簭�� �ð���ȸ (&F)"
               Height          =   345
               Left            =   1740
               TabIndex        =   8
               Top             =   2220
               Width           =   2295
            End
            Begin VB.TextBox txtFromLsnCD 
               Enabled         =   0   'False
               Height          =   300
               Left            =   690
               TabIndex        =   7
               Text            =   "txtFromLsnCD"
               Top             =   2250
               Width           =   915
            End
            Begin VB.ComboBox cboFromSubjCD 
               Height          =   300
               Left            =   690
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   3
               Top             =   840
               Width           =   1185
            End
            Begin VB.TextBox txtFromTcrCD 
               Height          =   300
               IMEMode         =   10  '�ѱ� 
               Left            =   1650
               TabIndex        =   2
               Text            =   "txtFromTcrCD"
               Top             =   480
               Width           =   1635
            End
            Begin VB.TextBox txtFromWeek 
               Height          =   300
               IMEMode         =   10  '�ѱ� 
               Left            =   690
               TabIndex        =   4
               Text            =   "txtFromWeek"
               Top             =   1200
               Width           =   915
            End
            Begin EditLib.fpLongInteger fpFromLesson 
               Height          =   300
               Left            =   690
               TabIndex        =   5
               Top             =   1560
               Width           =   915
               _Version        =   196608
               _ExtentX        =   1614
               _ExtentY        =   529
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
               MaxValue        =   "10"
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
            Begin EditLib.fpMask fpFromTcrCD 
               Height          =   300
               Left            =   690
               TabIndex        =   1
               Top             =   480
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   529
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
               Mask            =   "###"
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
            Begin EditLib.fpMask fpFromBan 
               Height          =   300
               Left            =   690
               TabIndex        =   6
               Top             =   1920
               Width           =   915
               _Version        =   196608
               _ExtentX        =   1614
               _ExtentY        =   529
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
               Mask            =   "AAA"
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
            Begin EditLib.fpMask fpYM 
               Height          =   285
               Left            =   690
               TabIndex        =   0
               Top             =   210
               Width           =   1005
               _Version        =   196608
               _ExtentX        =   1773
               _ExtentY        =   503
               Enabled         =   0   'False
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               ThreeDInsideStyle=   0
               ThreeDInsideHighlightColor=   -2147483633
               ThreeDInsideShadowColor=   -2147483642
               ThreeDInsideWidth=   1
               ThreeDOutsideStyle=   0
               ThreeDOutsideHighlightColor=   -2147483628
               ThreeDOutsideShadowColor=   -2147483632
               ThreeDOutsideWidth=   1
               ThreeDFrameWidth=   0
               BorderStyle     =   1
               BorderColor     =   16777215
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
               Appearance      =   1
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
            Begin VB.Label Label5 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   60
               TabIndex        =   36
               Top             =   525
               Width           =   465
            End
            Begin VB.Label Label6 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   60
               TabIndex        =   35
               Top             =   885
               Width           =   465
            End
            Begin VB.Label Label7 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   60
               TabIndex        =   34
               Top             =   1245
               Width           =   465
            End
            Begin VB.Label Label8 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   60
               TabIndex        =   33
               Top             =   1605
               Width           =   465
            End
            Begin VB.Label Label9 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "��"
               Height          =   210
               Left            =   60
               TabIndex        =   32
               Top             =   1965
               Width           =   465
            End
            Begin VB.Label Label10 
               BackStyle       =   0  '����
               Caption         =   "��)  �� : 1 ȭ : 2 ... �� : 6 �� : 7"
               Height          =   210
               Left            =   1800
               TabIndex        =   31
               Top             =   1245
               Width           =   2895
            End
            Begin VB.Label Label11 
               BackStyle       =   0  '����
               Caption         =   "��) 101  <- �迭( 1, 2 ) �� ( 01 )"
               Height          =   210
               Left            =   1860
               TabIndex        =   30
               Top             =   1965
               Width           =   2865
            End
         End
         Begin VB.Frame FraTo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "����Ǿ��� ����, �ð� �� �� ����"
            Height          =   2475
            Left            =   5220
            TabIndex        =   23
            Top             =   150
            Width           =   9105
            Begin FPSpread.vaSpread sprSEL 
               Height          =   2175
               Left            =   4860
               TabIndex        =   17
               Top             =   180
               Width           =   3885
               _Version        =   393216
               _ExtentX        =   6853
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
               MaxCols         =   11
               SpreadDesigner  =   "TMR052.frx":056B
            End
            Begin VB.TextBox txtToTcrCD 
               Height          =   300
               IMEMode         =   10  '�ѱ� 
               Left            =   1560
               TabIndex        =   10
               Text            =   "txtToTcrCD"
               Top             =   270
               Width           =   1635
            End
            Begin VB.ComboBox cboToSubjCD 
               Height          =   300
               Left            =   600
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   11
               Top             =   630
               Width           =   1185
            End
            Begin VB.TextBox txtToWeek 
               Height          =   300
               IMEMode         =   10  '�ѱ� 
               Left            =   600
               TabIndex        =   12
               Text            =   "txtToWeek"
               Top             =   990
               Width           =   915
            End
            Begin VB.TextBox txtToLsnCD 
               Enabled         =   0   'False
               Height          =   300
               Left            =   600
               TabIndex        =   15
               Text            =   "txtToLsnCD"
               Top             =   2040
               Width           =   915
            End
            Begin EditLib.fpLongInteger fpToLesson 
               Height          =   300
               Left            =   600
               TabIndex        =   13
               Top             =   1350
               Width           =   915
               _Version        =   196608
               _ExtentX        =   1614
               _ExtentY        =   529
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
               MaxValue        =   "10"
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
            Begin EditLib.fpMask fpToBan 
               Height          =   300
               Left            =   600
               TabIndex        =   14
               Top             =   1710
               Width           =   915
               _Version        =   196608
               _ExtentX        =   1614
               _ExtentY        =   529
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
               Mask            =   "AAA"
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
            Begin EditLib.fpMask fpToTcrCD 
               Height          =   300
               Left            =   600
               TabIndex        =   9
               Top             =   270
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
               _ExtentY        =   529
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
               Mask            =   "###"
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
            Begin VB.Label Label2 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   -30
               TabIndex        =   39
               Top             =   675
               Width           =   465
            End
            Begin VB.Label Label1 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   -30
               TabIndex        =   38
               Top             =   315
               Width           =   465
            End
            Begin VB.Label Label18 
               BackStyle       =   0  '����
               Caption         =   "��) 101  <- �迭( 1, 2 ) �� ( 01 )"
               Height          =   210
               Left            =   1770
               TabIndex        =   28
               Top             =   1755
               Width           =   2865
            End
            Begin VB.Label Label16 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "��"
               Height          =   210
               Left            =   -30
               TabIndex        =   27
               Top             =   1755
               Width           =   465
            End
            Begin VB.Label Label15 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   -30
               TabIndex        =   26
               Top             =   1395
               Width           =   465
            End
            Begin VB.Label Label14 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   -30
               TabIndex        =   25
               Top             =   1035
               Width           =   465
            End
            Begin VB.Label Label17 
               BackStyle       =   0  '����
               Caption         =   "��)  �� : 1 ȭ : 2 ... �� : 6 �� : 7"
               Height          =   210
               Left            =   1770
               TabIndex        =   24
               Top             =   1050
               Width           =   2865
            End
         End
         Begin FPSpread.vaSpread sprFromTCR 
            Height          =   2655
            Left            =   9030
            TabIndex        =   20
            Top             =   -1230
            Width           =   2955
            _Version        =   393216
            _ExtentX        =   5212
            _ExtentY        =   4683
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
            SpreadDesigner  =   "TMR052.frx":2123
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Height          =   945
            Left            =   210
            TabIndex        =   37
            Top             =   2790
            Width           =   4905
            Begin VB.CommandButton cmd_P_TmrNew 
               Caption         =   "�űԳ��� ����ϱ� (&N)"
               Height          =   555
               Left            =   60
               TabIndex        =   18
               Top             =   210
               Width           =   2265
            End
            Begin VB.CommandButton cmd_P_TmrDel 
               Caption         =   "��� �����ϱ� (&D)"
               Height          =   555
               Left            =   2490
               TabIndex        =   19
               Top             =   210
               Width           =   2265
            End
         End
         Begin FPSpread.vaSpread sprToTCR 
            Height          =   2655
            Left            =   12690
            TabIndex        =   40
            Top             =   -1440
            Width           =   2955
            _Version        =   393216
            _ExtentX        =   5212
            _ExtentY        =   4683
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
            SpreadDesigner  =   "TMR052.frx":3955
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   $"TMR052.frx":5187
            Height          =   1110
            Left            =   5550
            TabIndex        =   42
            Top             =   2700
            Width           =   4365
         End
      End
   End
End
Attribute VB_Name = "TMR052"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM026
'   �� ��  �� �� : �̵����� �ð�ǥ ���
'
'   ��   ��   �� : 2008/01/04
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private Sub lblClose_Click()
    Unload Me
End Sub







Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}", True
    
End Sub

Private Sub Form_Load()
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    
    Me.Top = (MDI001.Height - Me.Height) / 2
    Me.Left = (MDI001.Width - Me.Width) / 2
    Me.Width = 14910
    Me.Height = 13650
    
    Me.KeyPreview = True
    
    Me.Tag = "LOAD"
    
    fpYM.Text = Trim(TMR051.fpYM.UnFmtText)
    
    basFunction.RemoveContextMenu txtFromTcrCD
    basFunction.RemoveContextMenu txtToTcrCD
    
        With sprTmr
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            For nRow = 1 To .MaxRows Step 1
                For nCol = 1 To .MaxCols Step 1
                    .Row = nRow
                    .Col = nCol
                        Call basFunction.Set_SprType_Text(sprTmr, "top", "left", 100, " ")
                Next nCol
            Next nRow
            
        End With
        
        With sprSEL
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxRows = 0
            
        End With
    
    '> ������ ��
        With sprFromTCR
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
        End With
        
        With sprToTCR
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1

            .Tag = "0"
        End With
        
        fpFromTcrCD.Text = ""
        txtFromTcrCD.Text = ""
        
        fpToTcrCD.Text = ""
        txtToTcrCD.Text = ""
        
        cboFromSubjCD.Clear
        cboToSubjCD.Clear
        
        txtFromWeek.Text = ""
        fpFromLesson.Value = 1
        fpFromBan.Text = ""
        txtFromLsnCD.Text = ""
        
        sprFromTCR.Visible = False
        sprToTCR.Visible = False
        
    '> ����Ǿ��� ����,�ð� �� ��
        txtToWeek.Text = ""
        fpToLesson.Value = 1
        fpToBan.Text = ""
        txtToLsnCD.Text = ""
        
    Me.Tag = ""
    
End Sub






'############################################ ������ ���� #############################################################
'>> ������ȸ
Private Sub fpFromTcrCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyEscape
            sprFromTCR.Visible = False
            Exit Sub
        
        Case vbKeyBack
            txtFromTcrCD.Text = ""
            cboFromSubjCD.Clear
            Exit Sub
            
        Case vbKeyReturn, vbKeyTab
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRCD  LIKE '" & Trim(fpFromTcrCD.UnFmtText) & "%'"
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpFromTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtFromTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM"
            sStr = sStr & "   ORDER BY ACID, TCRCD "
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpFromTcrCD.Text = ""
            txtFromTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpFromTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpFromTcrCD.Text = Trim(.Fields("TCRCD"))
            txtFromTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtFromTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpFromTcrCD.Text) <> "" Then Call Find_From_TmrChg_Subj(fpFromTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprFromTCR.Visible = True
            sprFromTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprFromTCR.MaxRows = sprFromTCR.MaxRows + 1
                sprFromTCR.Row = sprFromTCR.MaxRows
                
                sprFromTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprFromTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprFromTCR.Top = fraFrom.Top + fpFromTcrCD.Top + fpFromTcrCD.Height
            sprFromTCR.Left = fraFrom.Left + fpFromTcrCD.Left
            sprFromTCR.Visible = True
            sprFromTCR.ZOrder 0
    
        End If
    End With
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'fpFromTcrCD.SetFocus
    cboFromSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub

Private Sub fpFromTcrCD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpFromTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtFromTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD, SUBJCD"
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpFromTcrCD.Text = ""
            txtFromTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpFromTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpFromTcrCD.Text = Trim(.Fields("TCRCD"))
            txtFromTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtFromTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpFromTcrCD.Text) <> "" Then Call Find_From_TmrChg_Subj(fpFromTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprFromTCR.Visible = True
            sprFromTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprFromTCR.MaxRows = sprFromTCR.MaxRows + 1
                sprFromTCR.Row = sprFromTCR.MaxRows
                
                sprFromTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprFromTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprFromTCR.Top = fraFrom.Top + fpFromTcrCD.Top + fpFromTcrCD.Height
            sprFromTCR.Left = fraFrom.Left + fpFromTcrCD.Left
            sprFromTCR.Visible = True
            sprFromTCR.ZOrder 0
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'fpFromTcrCD.SetFocus
    cboFromSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
    
End Sub



Private Sub txtFromTcrCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyBack
            fpFromTcrCD.Text = ""
            cboFromSubjCD.Clear
            
            Exit Sub
            
        Case vbKeyEscape
            sprFromTCR.Visible = False
            Exit Sub
                
        Case vbKeyReturn
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtFromTcrCD.Text) & "%'"
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtFromTcrCD.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtFromTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   ORDER BY ACID, TCRCD, SUBJCD"
        
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpFromTcrCD.Text = ""
            txtFromTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpFromTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpFromTcrCD.Text = Trim(.Fields("TCRCD"))
            txtFromTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtFromTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpFromTcrCD.Text) <> "" Then Call Find_From_TmrChg_Subj(fpFromTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprFromTCR.Visible = True
            sprFromTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprFromTCR.MaxRows = sprFromTCR.MaxRows + 1
                sprFromTCR.Row = sprFromTCR.MaxRows
                
                sprFromTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprFromTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprFromTCR.Top = fraFrom.Top + fpFromTcrCD.Top + fpFromTcrCD.Height
            sprFromTCR.Left = fraFrom.Left + fpFromTcrCD.Left
            sprFromTCR.Visible = True
            sprFromTCR.ZOrder 0
    
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'txtFromTcrCD.SetFocus
    cboFromSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub

Private Sub txtFromTcrCD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtFromTcrCD.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtFromTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
            
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpFromTcrCD.Text = ""
            txtFromTcrCD.Text = ""
        
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpFromTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpFromTcrCD.Text = Trim(.Fields("TCRCD"))
            txtFromTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtFromTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpFromTcrCD.Text) <> "" Then Call Find_From_TmrChg_Subj(fpFromTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprFromTCR.Visible = True
            sprFromTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprFromTCR.MaxRows = sprFromTCR.MaxRows + 1
                sprFromTCR.Row = sprFromTCR.MaxRows
                
                sprFromTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprFromTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprFromTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprFromTCR.Top = fraFrom.Top + fpFromTcrCD.Top + fpFromTcrCD.Height
            sprFromTCR.Left = fraFrom.Left + fpFromTcrCD.Left
            sprFromTCR.Visible = True
            sprFromTCR.ZOrder 0
    
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'txtFromTcrCD.SetFocus
    cboFromSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub







Private Sub sprFromTCR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            sprFromTCR.Visible = False
            
    End Select
End Sub

Private Sub sprFromTCR_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprFromTCR
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub

Private Sub sprFromTCR_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprFromTCR
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        '>> ������ �����ֱ�
        .Row = Row
        .Col = 1:       fpFromTcrCD.Text = Trim(.Text)
        .Col = 2:       txtFromTcrCD.Text = Trim(.Text)
        
        If Trim(fpFromTcrCD.Text) <> "" Then Call Find_From_TmrChg_Subj(fpFromTcrCD.Text)
        
        .Visible = False
        
        'fpFromTcrCD.SetFocus
        cboFromSubjCD.SetFocus
        
    End With
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
    sStr = sStr & "    SELECT LSNCD "
    sStr = sStr & "      From "
    
    sStr = sStr & "           (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 �߰�
    sStr = sStr & "              FROM SDLSN01TB "
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "            UNION"
    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "              FROM SDLSN02TB "
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "           )"
    
    sStr = sStr & "     WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "       AND KAEYOL  = '" & aKaeyol & "'"
    sStr = sStr & "       AND LSNCDNM = '" & aLsn & "'"
    sStr = sStr & "     GROUP BY LSNCD "
    
    
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        aLsnCD = ""
        
        If .RecordCount = 1 Then
            .MoveFirst
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
    MsgBox "���� �� ���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� �� ������ȸ"

End Sub


'## ������ȸ�� �ش簭���� ������ ��� ��ȸ
Private Sub Find_From_TmrChg_Subj(ByVal aTcr As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String

    Dim sSubjCD     As String
    Dim sSubjNM     As String

    On Error GoTo ErrStmt

    sStr = ""
    sStr = sStr & "  SELECT SUBJCD, SUBJNM"
    sStr = sStr & "    FROM SDTCR01TB"
    sStr = sStr & "   WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD = '" & Trim(aTcr) & "'"
    sStr = sStr & "   GROUP BY SUBJCD, SUBJNM "
    sStr = sStr & "   ORDER BY SUBJCD"

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        If .RecordCount = 0 Then
            
            cboFromSubjCD.Clear
            cboFromSubjCD.AddItem "����" & Space(30) & "X"
            
        Else
            cboFromSubjCD.Clear
            
            .MoveFirst

            For nRec = 1 To .RecordCount Step 1

                sSubjCD = ""
                sSubjNM = ""

                If IsNull(.Fields("SUBJCD")) = False Then sSubjCD = Trim(.Fields("SUBJCD"))
                If IsNull(.Fields("SUBJNM")) = False Then sSubjNM = Trim(.Fields("SUBJNM"))

                cboFromSubjCD.AddItem sSubjNM & Space(30) & sSubjCD

                .MoveNext
            Next nRec
        End If
    End With

    If cboFromSubjCD.ListCount > 0 Then cboFromSubjCD.ListIndex = 0

    Set DBCmd = Nothing
    Set DBRec = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    On Error GoTo 0
    MsgBox "������ ���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� ������ȸ"
End Sub





'## ����, ����, ����, ���ÿ� �ش��ϴ� ���� ��ȸ
Private Sub fpFromBan_GotFocus()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String
    Dim sWeek       As String

    On Error GoTo ErrStmt

    If Trim(fpFromTcrCD.UnFmtText) = "" Then Exit Sub
    If Trim(Right(cboFromSubjCD.Text, 30)) = "" Or _
       Trim(Right(cboFromSubjCD.Text, 30)) = "X" Then Exit Sub
    If Trim(txtFromWeek.Text) = "" Then Exit Sub
    If fpFromLesson.Value < 1 Or fpFromLesson.Value > 10 Then Exit Sub

    Select Case Trim(txtFromWeek.Text)
        Case "1"
            sWeek = "2"
        Case "2"
            sWeek = "3"
        Case "3"
            sWeek = "4"
        Case "4"
            sWeek = "5"
        Case "5"
            sWeek = "6"
        Case "6"
            sWeek = "7"
        Case "7"
            sWeek = "1"
    End Select

    sStr = ""
    sStr = sStr & "  SELECT A.LSNCD, "
    
    Select Case Trim(basModule.SchCD)
        Case "N"
            sStr = sStr & " SUBSTR(B.KAEYOL,2,1)||LSNCDNM AS BAN"
        Case "S"
            sStr = sStr & " SUBSTR(B.KAEYOL,2,1)||LSNCDNM AS BAN"
        Case "K"
            sStr = sStr & " SUBSTR(GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD), 1, 1)||B.LSNCDNM AS BAN "
    End Select
    
    sStr = sStr & "    FROM SDTRX50TB A, "
    
    sStr = sStr & "         (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 �߰�
    sStr = sStr & "            FROM SDLSN01TB "
    sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "          UNION"
    sStr = sStr & "          SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "            FROM SDLSN02TB "
    sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         ) B "
    
    sStr = sStr & "   Where A.ACID  = B.ACID"
    sStr = sStr & "     AND A.LSNCD = B.LSNCD"
    sStr = sStr & "     AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "     AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD   = '" & Trim(fpFromTcrCD.UnFmtText) & "'"
    sStr = sStr & "     AND SUBJCD  = '" & Trim(Right(cboFromSubjCD.Text, 30)) & "'"
    sStr = sStr & "     AND WEEKS   = " & sWeek
    sStr = sStr & "     AND LESSON  = " & Trim(CStr(fpFromLesson.UnFmtText))

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        If .RecordCount = 0 Then
            txtFromLsnCD.Text = ""
            fpFromBan.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            txtFromLsnCD.Text = ""
                If IsNull(.Fields("LSNCD")) = False Then
                    txtFromLsnCD.Text = Trim(.Fields("LSNCD"))
                    txtToLsnCD.Text = Trim(.Fields("LSNCD"))        '< �⺻��
                End If
            fpFromBan.Text = ""
                If IsNull(.Fields("BAN")) = False Then
                    fpFromBan.Text = Trim(.Fields("BAN"))
                    fpToBan.Text = Trim(.Fields("BAN"))             '< �⺻��
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
    MsgBox "������ �ð�ǥ ��ϳ��� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ����"


End Sub






'## 101 <- �迭 �� ���ڵ��Ī���� ���� �˻�.
Private Sub fpToBan_LostFocus()
    Dim sLsnCD      As String
    Dim sLsn        As String
    Dim sKaeyol     As String
    
    
    If Trim(fpToBan.UnFmtText) = "" Then Exit Sub
    If Len(fpToBan.UnFmtText) <> 3 Then Exit Sub
    
    sKaeyol = "0" & Left(Trim(fpToBan.UnFmtText), 1)
    sLsn = Right(Trim(fpToBan.UnFmtText), 2)
    
    sLsnCD = ""
    Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
    
    If Len(sLsnCD) = 5 Then
        txtToLsnCD.Text = sLsnCD
        
    End If
    
End Sub


'## ��ϳ��� ����
Private Sub cmd_P_TmrDel_Click()
        Dim sKaeyol     As String
    Dim sLsn        As String
    Dim sLsnCD      As String
    
    Dim sF_TcrCD    As String
    Dim sF_TcrNM    As String
    Dim sF_SubjCD   As String
    Dim sF_SubjNM   As String
    Dim sF_LsnCD    As String
    Dim sF_LsnCDNM  As String
    Dim sF_Weeks    As String
    Dim sF_Lesson   As String
    
    Dim bRet        As Boolean
    Dim nRow        As Long
    Dim nCol        As Long
    Dim sComp       As String
    
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    '>> �ð�ǥ �ڵ� üũ
        If Trim(fpYM.UnFmtText) = "" Then
            MsgBox "�ð�ǥ �ڵ带 Ȯ���ϼ���.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
    
    '>> 1. �ݳ��� -> ���ڵ�� �ٲٱ� ( ������ �ð�ǥ ���� )
        If Trim(fpFromBan.UnFmtText) = "" Or Len(fpFromBan.UnFmtText) <> 3 Then
            MsgBox "������ �ð�ǥ �������� �� ������ ��������.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
        
        Select Case Trim(basModule.SchCD)
            Case "K"
            
                Select Case Left(Trim(fpFromBan.UnFmtText), 1)
                    Case "1", "3", "5"
                        sKaeyol = "01"          ' ���� �ι���
                    Case "2", "4", "6"
                        sKaeyol = "02"          ' ���� �ڿ���
                End Select
            
            Case Else
                sKaeyol = "0" & Left(Trim(fpFromBan.UnFmtText), 1)
        End Select
        
        sLsn = Right(Trim(fpFromBan.UnFmtText), 2)
        
        Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
        
        If Trim(sLsnCD) = "" Then
            MsgBox "������ �ð�ǥ �������� �� ������ �ش��ϴ� ������ ������ Ȯ���Ͻʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        Else
            txtFromLsnCD.Text = Trim(sLsnCD)
        End If
    
    ' ����üũ
        If Trim(fpFromTcrCD.UnFmtText) = "" Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Len(fpFromTcrCD.UnFmtText) <> 3 Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Trim(Right(cboFromSubjCD.Text, 30)) = "X" Then
            MsgBox "������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub          '< �������.
        End If
        
        Select Case Trim(txtFromWeek.Text)
            Case "1"
                sF_Weeks = "2"
            Case "2"
                sF_Weeks = "3"
            Case "3"
                sF_Weeks = "4"
            Case "4"
                sF_Weeks = "5"
            Case "5"
                sF_Weeks = "6"
            Case "6"
                sF_Weeks = "7"
            Case "7"
                sF_Weeks = "1"
            Case Else
                MsgBox "������ Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        
        Select Case CLng(fpFromLesson.Text)
            Case 1 To 10
                sF_Lesson = Trim(fpFromLesson.Text)
            Case Else
                MsgBox "���ø� Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        If Trim(txtFromLsnCD.Text) = "" Then
            MsgBox "������ ������ ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
     
    '** �����׸� �����ϱ�
        sF_TcrCD = Trim(fpFromTcrCD.UnFmtText)
        sF_TcrNM = Trim(txtFromTcrCD.Text)
        sF_SubjCD = Trim(Right(cboFromSubjCD.Text, 30))
        sF_SubjNM = Trim(Mid(cboFromSubjCD.Text, 1, Len(cboFromSubjCD.Text) - 10))
        sF_LsnCD = Trim(txtFromLsnCD.Text)
        'sF_Weeks       <- �̹� ������ ó��
        'sF_Lesson      <- �̹� ������ ó��
        
        
        bRet = Del_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson)
        If bRet = True Then
            ' ����,���� & �� ���� ����
            For nRow = 1 To TMR051.sprTmr_Lsn.MaxRows Step 1
                TMR051.sprTmr_Lsn.Row = nRow
                TMR051.sprTmr_Lsn.Col = SpreadHeader + 1        '< ����
                
                If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_Weeks, vbTextCompare) = 0 Then
                    nr_Chk = TMR051.sprTmr_Lsn.Row              '< row ��

                    TMR051.sprTmr_Lsn.Col = SpreadHeader + 2        '< lesson
                    
                    If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_Lesson, vbTextCompare) = 0 Then
                        
                        For nCol = 1 To TMR051.sprTmr_Lsn.MaxCols Step 1
                            TMR051.sprTmr_Lsn.Col = nCol
                            TMR051.sprTmr_Lsn.Row = SpreadHeader + 1
                        
                            If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_LsnCD, vbTextCompare) = 0 Then
                                nc_Chk = TMR051.sprTmr_Lsn.Col
                                
                                TMR051.sprTmr_Lsn.Row = nr_Chk
                                TMR051.sprTmr_Lsn.Col = nc_Chk
                                    TMR051.sprTmr_Lsn.Text = ""
                                    
                                Exit For
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
            ' ���� & ���� ���� ����
            For nRow = 1 To TMR051.sprTmr_Tcr.MaxRows Step 1
                TMR051.sprTmr_Tcr.Row = nRow
                TMR051.sprTmr_Tcr.Col = SpreadHeader
                
                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_TcrCD, vbTextCompare) = 0 Then
                    TMR051.sprTmr_Tcr.Col = SpreadHeader + 1
                    
                    If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_SubjCD, vbTextCompare) = 0 Then
                        nr_Chk = TMR051.sprTmr_Tcr.Row
                        
                        For nCol = 1 To TMR051.sprTmr_Tcr.MaxCols Step 1
                            TMR051.sprTmr_Tcr.Col = nCol
                            TMR051.sprTmr_Tcr.Row = SpreadHeader + 1
                            
                            If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_Weeks, vbTextCompare) = 0 Then
                                TMR051.sprTmr_Tcr.Row = SpreadHeader + 2
                                
                                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_Lesson, vbTextCompare) = 0 Then
                                    nc_Chk = TMR051.sprTmr_Tcr.Col
                                    
                                    TMR051.sprTmr_Tcr.Row = nr_Chk
                                    TMR051.sprTmr_Tcr.Col = nc_Chk
                                        TMR051.sprTmr_Tcr.Text = ""
                                        
                                    Exit For
                                End If
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
        End If
        
        
    Call Save_Log_Chg_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson, _
                               "", "", "DEL", "", "")
    
    '<< �ʱ�ȭ >>
'    fpFromTcrCD.Text = ""
'    txtFromTcrCD.Text = ""
'    cboFromSubjCD.Clear
'    txtFromWeek.Text = ""
'    fpFromLesson.Value = 1
'    fpFromBan.Text = ""
'    txtFromLsnCD.Text = ""
    
    sprFromTCR.Visible = False
        
    '> ����Ǿ��� ����,�ð� �� ��
'    txtToWeek.Text = ""
'    fpToLesson.Value = 1
'    fpToBan.Text = ""
'    txtToLsnCD.Text = ""
    
    fpFromTcrCD.SetFocus
    
    MsgBox "�����Ͽ����ϴ�." & vbCrLf & _
           "Ȯ���ϼ���", vbInformation + vbOKOnly, "�ð�ǥ �����ϱ�"
        
End Sub

'## �űԵ��
Private Sub cmd_P_TmrNew_Click()
    Dim sKaeyol     As String
    Dim sLsn        As String
    Dim sLsnCD      As String
    
    Dim sF_TcrCD    As String
    Dim sF_TcrNM    As String
    Dim sF_SubjCD   As String
    Dim sF_SubjNM   As String
    Dim sF_LsnCD    As String
    Dim sF_LsnCDNM  As String
    Dim sF_Weeks    As String
    Dim sF_Lesson   As String
    
    Dim bRet        As Boolean
    Dim nRow        As Long
    Dim nCol        As Long
    Dim sComp       As String
    
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    '>> �ð�ǥ �ڵ� üũ
        If Trim(fpYM.UnFmtText) = "" Then
            MsgBox "�ð�ǥ �ڵ带 Ȯ���ϼ���.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
    
    '>> 1. �ݳ��� -> ���ڵ�� �ٲٱ� ( ������ �ð�ǥ ���� )
        If Trim(fpFromBan.UnFmtText) = "" Or Len(fpFromBan.UnFmtText) <> 3 Then
            MsgBox "������ �ð�ǥ �������� �� ������ ��������.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
        
        Select Case Trim(basModule.SchCD)
            Case "K"
            
                Select Case Left(Trim(fpFromBan.UnFmtText), 1)
                    Case "1", "3", "5"
                        sKaeyol = "01"          ' ���� �ι���
                    Case "2", "4", "6"
                        sKaeyol = "02"          ' ���� �ڿ���
                End Select
            
            Case Else
                sKaeyol = "0" & Left(Trim(fpFromBan.UnFmtText), 1)
        End Select
        
        sLsn = Right(Trim(fpFromBan.UnFmtText), 2)
        
        Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
        
        If Trim(sLsnCD) = "" Then
            MsgBox "������ �ð�ǥ �������� �� ������ �ش��ϴ� ������ ������ Ȯ���Ͻʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        Else
            txtFromLsnCD.Text = Trim(sLsnCD)
        End If
    
    ' ����üũ
        If Trim(fpFromTcrCD.UnFmtText) = "" Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Len(fpFromTcrCD.UnFmtText) <> 3 Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Trim(Right(cboFromSubjCD.Text, 30)) = "X" Then
            MsgBox "������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub          '< �������.
        End If
        
        Select Case Trim(txtFromWeek.Text)
            Case "1"
                sF_Weeks = "2"
            Case "2"
                sF_Weeks = "3"
            Case "3"
                sF_Weeks = "4"
            Case "4"
                sF_Weeks = "5"
            Case "5"
                sF_Weeks = "6"
            Case "6"
                sF_Weeks = "7"
            Case "7"
                sF_Weeks = "1"
            Case Else
                MsgBox "������ Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        
        Select Case CLng(fpFromLesson.Text)
            Case 1 To 10
                sF_Lesson = Trim(fpFromLesson.Text)
            Case Else
                MsgBox "���ø� Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        If Trim(txtFromLsnCD.Text) = "" Then
            MsgBox "������ ������ ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
     
    '** �����׸� �����ϱ�
        sF_TcrCD = Trim(fpFromTcrCD.UnFmtText)
        sF_TcrNM = Trim(txtFromTcrCD.Text)
        sF_SubjCD = Trim(Right(cboFromSubjCD.Text, 30))
        sF_SubjNM = Trim(Mid(cboFromSubjCD.Text, 1, Len(cboFromSubjCD.Text) - 10))
        sF_LsnCD = Trim(txtFromLsnCD.Text)
        'sF_Weeks       <- �̹� ������ ó��
        'sF_Lesson      <- �̹� ������ ó��
        
        bRet = Save_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson)
        If bRet = True Then
            ' ����,���� & �� ���� ���
            For nRow = 1 To TMR051.sprTmr_Lsn.MaxRows Step 1
                TMR051.sprTmr_Lsn.Row = nRow
                TMR051.sprTmr_Lsn.Col = SpreadHeader + 1        '< ����
                
                If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_Weeks, vbTextCompare) = 0 Then
                    nr_Chk = TMR051.sprTmr_Lsn.Row              '< row ��

                    TMR051.sprTmr_Lsn.Col = SpreadHeader + 2        '< lesson
                    
                    If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_Lesson, vbTextCompare) = 0 Then
                        
                        For nCol = 1 To TMR051.sprTmr_Lsn.MaxCols Step 1
                            TMR051.sprTmr_Lsn.Col = nCol
                            TMR051.sprTmr_Lsn.Row = SpreadHeader + 1
                        
                            If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sF_LsnCD, vbTextCompare) = 0 Then
                                nc_Chk = TMR051.sprTmr_Lsn.Col
                                
                                TMR051.sprTmr_Lsn.Row = nr_Chk
                                TMR051.sprTmr_Lsn.Col = nc_Chk
                                    TMR051.sprTmr_Lsn.Text = sF_SubjNM & "," & sF_TcrNM
                                    
                                Exit For
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
            ' ���� & ���� ���� ���
            For nRow = 1 To TMR051.sprTmr_Tcr.MaxRows Step 1
                TMR051.sprTmr_Tcr.Row = nRow
                TMR051.sprTmr_Tcr.Col = SpreadHeader
                
                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_TcrCD, vbTextCompare) = 0 Then
                    TMR051.sprTmr_Tcr.Col = SpreadHeader + 1
                    
                    If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_SubjCD, vbTextCompare) = 0 Then
                        nr_Chk = TMR051.sprTmr_Tcr.Row
                        
                        For nCol = 1 To TMR051.sprTmr_Tcr.MaxCols Step 1
                            TMR051.sprTmr_Tcr.Col = nCol
                            TMR051.sprTmr_Tcr.Row = SpreadHeader + 1
                            
                            If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_Weeks, vbTextCompare) = 0 Then
                                TMR051.sprTmr_Tcr.Row = SpreadHeader + 2
                                
                                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_Lesson, vbTextCompare) = 0 Then
                                    nc_Chk = TMR051.sprTmr_Tcr.Col
                                    
                                    TMR051.sprTmr_Tcr.Row = nr_Chk
                                    TMR051.sprTmr_Tcr.Col = nc_Chk
                                        TMR051.sprTmr_Tcr.Text = sF_LsnCDNM
                                        
                                    Exit For
                                End If
                            End If
                        Next nCol
                    End If
                End If
            Next nRow
            
        End If
        
        
    Call Save_Log_Chg_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson, _
                               "", "", "SAVE", "", "")
    
    '<< �ʱ�ȭ >>
    'fpFromTcrCD.Text = ""
    'txtFromTcrCD.Text = ""
    'cboFromSubjCD.Clear
    'txtFromWeek.Text = ""
    'fpFromLesson.Value = 1
    'fpFromBan.Text = ""
    'txtFromLsnCD.Text = ""
    
    sprFromTCR.Visible = False
        
    '> ����Ǿ��� ����,�ð� �� ��
'    txtToWeek.Text = ""
'    fpToLesson.Value = 1
'    fpToBan.Text = ""
'    txtToLsnCD.Text = ""
    
    fpFromTcrCD.SetFocus
    
    MsgBox "����Ͽ����ϴ�." & vbCrLf & _
           "Ȯ���ϼ���", vbInformation + vbOKOnly, "�ð�ǥ �����ϱ�"
    
End Sub

'ó��
Private Sub cmd_P_TmrChg_Click()
    Dim sLsnCD      As String
    Dim sKaeyol     As String
    Dim sLsn        As String
    
    Dim sTmp        As String
    
    Dim sF_TcrCD    As String
    Dim sF_TcrNM    As String
    Dim sF_SubjCD   As String
    Dim sF_SubjNM   As String
    Dim sF_LsnCD    As String
    Dim sF_Lesson   As String
    Dim sF_Weeks    As String
    
    Dim sT_TcrCD    As String
    Dim sT_SubjCD   As String
    Dim sT_LsnCD    As String
    Dim sT_LsnCDNM  As String
    Dim sT_Lesson   As String
    Dim sT_Weeks    As String
    
    Dim bRet        As Boolean
    Dim nRow        As Long
    Dim nCol        As Long
    Dim sComp       As String
    
    Dim nr_Chk      As Long
    Dim nc_Chk      As Long
    
    '>> �ð�ǥ �ڵ� üũ
        If Trim(fpYM.UnFmtText) = "" Then
            MsgBox "�ð�ǥ �ڵ带 Ȯ���ϼ���.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
    
    '>> 1. �ݳ��� -> ���ڵ�� �ٲٱ� ( ������ �ð�ǥ ���� )
        If Trim(fpFromBan.UnFmtText) = "" Or Len(fpFromBan.UnFmtText) <> 3 Then
            MsgBox "������ �ð�ǥ �������� �� ������ ��������.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
        
        Select Case Trim(basModule.SchCD)
            Case "K"
            
                Select Case Left(Trim(fpFromBan.UnFmtText), 1)
                    Case "1", "3", "5"
                        sKaeyol = "01"          ' ���� �ι���
                    Case "2", "4", "6"
                        sKaeyol = "02"          ' ���� �ڿ���
                End Select
            
            Case Else
                sKaeyol = "0" & Left(Trim(fpFromBan.UnFmtText), 1)
        End Select
        
        sLsn = Right(Trim(fpFromBan.UnFmtText), 2)
        
        Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
        
        If Trim(sLsnCD) = "" Then
            MsgBox "������ �ð�ǥ �������� �� ������ �ش��ϴ� ������ ������ Ȯ���Ͻʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
        
    '>> 2. �ݳ��� -> ���ڵ�� �ٲٱ� (����Ǿ��� �ð�ǥ ����) <== LOSTFOCUS�ÿ� �߻�
        
        sLsnCD = ""
        sKaeyol = ""
        sLsn = ""
        
        If Trim(fpToBan.UnFmtText) = "" Then Exit Sub
        If Len(fpToBan.UnFmtText) <> 3 Then Exit Sub
        
        sKaeyol = "0" & Left(Trim(fpToBan.UnFmtText), 1)
        sLsn = Right(Trim(fpToBan.UnFmtText), 2)
        
        sLsnCD = ""
        Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
    
        If Len(sLsnCD) = 5 Then
            txtToLsnCD.Text = sLsnCD
            
        End If
        
    ' ����üũ
        If Trim(fpFromTcrCD.UnFmtText) = "" Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Len(fpFromTcrCD.UnFmtText) <> 3 Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Trim(Right(cboFromSubjCD.Text, 30)) = "X" Then
            MsgBox "������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub          '< �������.
        End If
        
        Select Case Trim(txtFromWeek.Text)
            Case "1"
                sF_Weeks = "2"
            Case "2"
                sF_Weeks = "3"
            Case "3"
                sF_Weeks = "4"
            Case "4"
                sF_Weeks = "5"
            Case "5"
                sF_Weeks = "6"
            Case "6"
                sF_Weeks = "7"
            Case "7"
                sF_Weeks = "1"
            Case Else
                MsgBox "������ Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        
        Select Case CLng(fpFromLesson.Text)
            Case 1 To 10
                sF_Lesson = Trim(fpFromLesson.Text)
            Case Else
                MsgBox "���ø� Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        If Trim(txtFromLsnCD.Text) = "" Then
            MsgBox "������ ������ ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        
        '>> ����Ǿ��� ******
        If Trim(fpToTcrCD.UnFmtText) = "" Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Len(fpToTcrCD.UnFmtText) <> 3 Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Trim(Right(cboToSubjCD.Text, 30)) = "X" Then
            MsgBox "������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub          '< �������.
        End If
        
        Select Case Trim(txtToWeek.Text)
            Case "1"
                sT_Weeks = "2"
            Case "2"
                sT_Weeks = "3"
            Case "3"
                sT_Weeks = "4"
            Case "4"
                sT_Weeks = "5"
            Case "5"
                sT_Weeks = "6"
            Case "6"
                sT_Weeks = "7"
            Case "7"
                sT_Weeks = "1"
            Case Else
                MsgBox "������ Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        
        Select Case CLng(fpToLesson.Text)
            Case 1 To 10
                sT_Lesson = Trim(fpToLesson.Text)
            Case Else
                MsgBox "���ø� Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        If Trim(txtToLsnCD.Text) = "" Then
            MsgBox "������ ������ ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        
        
    '** �����ؾߵǴ� ���� ���� ( A -> B ���� A ) **
        sF_TcrCD = Trim(fpFromTcrCD.UnFmtText)
        sF_TcrNM = Trim(txtFromTcrCD.Text)
        sF_SubjCD = Trim(Right(cboFromSubjCD.Text, 30))
        sF_SubjNM = Trim(Mid(cboFromSubjCD.Text, 1, Len(cboFromSubjCD.Text) - 10))
        sF_LsnCD = Trim(txtFromLsnCD.Text)
        'sF_Weeks       <- �̹� ������ ó��
        'sF_Lesson      <- �̹� ������ ó��
        
        bRet = Del_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson)         '<- �������� ����
        If bRet = True Then
            ' ����,���� & �� ���� ����
            Call Show_TMR_Tcr(sF_Weeks, sF_Lesson)
            
            ' ���� & ���� ���� ����
            Call Show_TMR_Lsn(sF_Weeks, sF_Lesson)
            
        End If
        
        
    '** �����׸� �����ϱ�
        
        sT_TcrCD = Trim(txtToTcrCD.Text)
        sT_SubjCD = Trim(Right(cboToSubjCD.Text, 30))
        sT_LsnCD = Trim(txtToLsnCD.Text)
        sT_LsnCDNM = Trim(fpToBan.UnFmtText)
        'ST_Weeks       <- �̹� ������ ó��
        'ST_Lesson      <- �̹� ������ ó��
            
        bRet = Save_TMR_Data(sF_TcrCD, sF_SubjCD, sT_LsnCD, sT_Weeks, sT_Lesson)
        If bRet = True Then
            ' ����,���� & �� ���� ����
            Call Show_TMR_Tcr(sF_Weeks, sF_Lesson)
            
            ' ���� & ���� ���� ����
            Call Show_TMR_Lsn(sF_Weeks, sF_Lesson)
            
        End If
        
        
    Call Save_Log_Chg_TMR_Data(sF_TcrCD, sF_SubjCD, sF_LsnCD, sF_Weeks, sF_Lesson, _
                               sT_TcrCD, sT_SubjCD, sT_LsnCD, sT_Weeks, sT_Lesson)
    
    '<< �ʱ�ȭ >>
'    fpFromTcrCD.Text = ""
'    txtFromTcrCD.Text = ""
'    cboFromSubjCD.Clear
'    txtFromWeek.Text = ""
'    fpFromLesson.Value = 1
'    fpFromBan.Text = ""
'    txtFromLsnCD.Text = ""
    
    sprFromTCR.Visible = False
        
    '> ����Ǿ��� ����,�ð� �� ��
    'fpToTcrCD.Text = ""
    'txtToTcrCD.Text = ""
    'cboToSubjCD.Clear
    'txtToWeek.Text = ""
    'fpToLesson.Value = 1
    'fpToBan.Text = ""
    'txtToLsnCD.Text = ""
        
    fpFromTcrCD.SetFocus
    
    MsgBox "ó���Ͽ����ϴ�." & vbCrLf & _
           "Ȯ���ϼ���", vbInformation + vbOKOnly, "�ð�ǥ �����ϱ�"
    
End Sub










'###############################################################################################################################################################


'## ��ü �ð�ǥ �������� �����ֱ� : TMR051.sprTmr_Lsn
Public Sub Show_TMR_Lsn(ByVal aWeek As String, ByVal aLesson As String)
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
    sStr = sStr & " SELECT GET_SUBJNM(ACID, TCRCD, SUBJCD)||','||GET_TCRNM(ACID, TCRCD) AS DS, LSNCD, WEEKS, LESSON"
    sStr = sStr & "   From SDTRX50TB"
    sStr = sStr & "  WHERE YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND WEEKS  = " & aWeek
    sStr = sStr & "    AND LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With TMR051.sprTmr_Lsn
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
                    
                    If .BackColor = basModule.SectionColor1 Or _
                       .BackColor = TMR051.lblNotTeaching.BackColor Then
                        ' no action
                    Else
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.WhiteColor
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End If
                    
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
            
            With TMR051.sprTmr_Lsn
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
                                        Call basFunction.Set_SprType_Text(TMR051.sprTmr_Lsn, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(TMR051.sprTmr_Lsn, "center", "left", 60, sData)
                                        
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






'## ��ü �ð�ǥ �������� �����ֱ� : TMR051.sprTmr_Tcr
Public Sub Show_TMR_Tcr(ByVal aWeek As String, ByVal aLesson As String)
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
    sStr = sStr & " SELECT TCRCD, SUBJCD, GET_KEAYOL_N_LSN_TCR01(ACID, LSNCD) AS DS"
    sStr = sStr & "   From SDTRX50TB"
    sStr = sStr & "  WHERE YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "    AND ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND WEEKS  = " & aWeek
    sStr = sStr & "    AND LESSON = " & aLesson
        
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    
    With TMR051.sprTmr_Tcr
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
                    
                    If .BackColor = basModule.SectionColor1 Or _
                       .BackColor = TMR051.lblNotTeaching.BackColor Then
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
            
            With TMR051.sprTmr_Tcr
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
                                        Call basFunction.Set_SprType_Text(TMR051.sprTmr_Tcr, "center", "left", 60, sData)
                                    End If
                                Else
                                    If InStr(1, Trim(.Text), sData, vbTextCompare) = 0 Then
                                        sData = sData & "/" & Trim(.Text)
                                        Call basFunction.Set_SprType_Text(TMR051.sprTmr_Tcr, "center", "left", 60, sData)
                                        
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
































'###############################################################################################################################################################















































'## ���泻�� LOG ������ ����� << insert ��
Private Sub Save_Log_Chg_TMR_Data(ByVal aF_TcrCD As String, ByVal aF_SubjCD As String, ByVal aF_LsnCD As String, ByVal aF_Weeks As String, ByVal aF_Lesson As String, _
                                  ByVal aT_TcrCD As String, ByVal aT_SubjCD As String, ByVal aT_LsnCD As String, ByVal aT_Weeks As String, ByVal aT_Lesson As String)

    
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim nExe        As Long
    
    Dim ni          As Integer
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX52TB( YM      , ACID    , "
    sStr = sStr & "                         F_TCRCD , F_SUBJCD, F_LSNCD , F_LESSON, F_WEEKS , "
    sStr = sStr & "                         T_TCRCD , T_SUBJCD, T_LSNCD , T_LESSON, T_WEEKS  ) "
    sStr = sStr & "  VALUES ( "
    sStr = sStr & "           '" & Trim(fpYM.UnFmtText) & "', "
    sStr = sStr & "           '" & Trim(basModule.SchCD) & "', "
    sStr = sStr & "           '" & Trim(aF_TcrCD) & "', '" & Trim(aF_SubjCD) & "', '" & Trim(aF_LsnCD) & "', " & Trim(aF_Lesson) & ", " & Trim(aF_Weeks) & ", "
    sStr = sStr & "           '" & Trim(aT_TcrCD) & "', '" & Trim(aT_SubjCD) & "', '" & Trim(aT_LsnCD) & "', " & Trim(aT_Lesson) & ", " & Trim(aT_Weeks)
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
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
    Else
        
ErrStmt:
        basDataBase.DBConn.RollbackTrans
        Set DBCmd = Nothing
        Set DBParam = Nothing
    End If
    
End Sub


'## �ð�ǥ ����
Private Function Save_TMR_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, _
                               ByVal aLsnCD As String, ByVal aWeeks As String, ByVal aLesson As String) As Boolean

    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim nExe        As Long
    
    Dim ni          As Integer
    Dim bRet        As Boolean
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX50TB(YM, ACID, TCRCD, SUBJCD, LSNCD, LESSON, WEEKS ) "
    sStr = sStr & "  VALUES ( "
    sStr = sStr & "           '" & Trim(fpYM.UnFmtText) & "', "
    sStr = sStr & "           '" & Trim(basModule.SchCD) & "', "
    sStr = sStr & "           '" & Trim(aTcrCD) & "', "
    sStr = sStr & "           '" & Trim(aSubjCD) & "', "
    sStr = sStr & "           '" & Trim(aLsnCD) & "', "
    sStr = sStr & "           " & Trim(aLesson) & ", "
    sStr = sStr & "           " & Trim(aWeeks)
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
        bRet = True
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Save_TMR_Data = bRet
        
    Else
        
ErrStmt:
        basDataBase.DBConn.RollbackTrans
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Save_TMR_Data = bRet
        
    End If
    
End Function

'## �ð�ǥ ���� ����
Private Function Del_TMR_Data(ByVal aTcrCD As String, ByVal aSubjCD As String, _
                              ByVal aLsnCD As String, ByVal aWeeks As String, ByVal aLesson As String) As Boolean

    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sTmp        As String
    Dim nExe        As Long
    
    Dim ni          As Integer
    Dim bRet        As Boolean
    
    On Error GoTo ErrStmt
    bRet = False
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


    
    sStr = ""
    sStr = sStr & "  DELETE SDTRX50TB"
    sStr = sStr & "   WHERE YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "     AND ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD  = '" & Trim(aTcrCD) & "'"
    sStr = sStr & "     AND SUBJCD = '" & Trim(aSubjCD) & "'"
    sStr = sStr & "     AND LSNCD  = '" & Trim(aLsnCD) & "'"
    sStr = sStr & "     AND LESSON = " & Trim(aLesson)
    sStr = sStr & "     AND WEEKS  = " & Trim(aWeeks)
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
                    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
            
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Del_TMR_Data = bRet
    Else
        
ErrStmt:
        basDataBase.DBConn.RollbackTrans
        Set DBCmd = Nothing
        Set DBParam = Nothing
        
        Del_TMR_Data = bRet
        
    End If
    
End Function



























'#############################################################################################################################################################
'>> ������ȸ �޴� �� �κ�
'#############################################################################################################################################################
Private Sub fpToTcrCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyEscape
            sprToTCR.Visible = False
            Exit Sub
        
        Case vbKeyBack
            txtToTcrCD.Text = ""
            cboToSubjCD.Clear
            Exit Sub
            
        Case vbKeyReturn, vbKeyTab
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRCD  LIKE '" & Trim(fpToTcrCD.UnFmtText) & "%'"
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpToTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtToTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM"
            sStr = sStr & "   ORDER BY ACID, TCRCD "
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpToTcrCD.Text = ""
            txtToTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpToTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpToTcrCD.Text = Trim(.Fields("TCRCD"))
            txtToTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtToTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpToTcrCD.Text) <> "" Then Call Find_To_TmrChg_Subj(fpToTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprToTCR.Visible = True
            sprToTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprToTCR.MaxRows = sprToTCR.MaxRows + 1
                sprToTCR.Row = sprToTCR.MaxRows
                
                sprToTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprToTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprToTCR.Top = FraTo.Top + fpToTcrCD.Top + fpToTcrCD.Height
            sprToTCR.Left = FraTo.Left + fpToTcrCD.Left
            sprToTCR.Visible = True
            sprToTCR.ZOrder 0
    
        End If
    End With
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'fpToTcrCD.SetFocus
    cboToSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub

Private Sub fpToTcrCD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpToTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtToTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD"
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpToTcrCD.Text = ""
            txtToTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpToTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpToTcrCD.Text = Trim(.Fields("TCRCD"))
            txtToTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtToTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpToTcrCD.Text) <> "" Then Call Find_To_TmrChg_Subj(fpToTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprToTCR.Visible = True
            sprToTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprToTCR.MaxRows = sprToTCR.MaxRows + 1
                sprToTCR.Row = sprToTCR.MaxRows
                
                sprToTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprToTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprToTCR.Top = FraTo.Top + fpToTcrCD.Top + fpToTcrCD.Height
            sprToTCR.Left = FraTo.Left + fpToTcrCD.Left
            sprToTCR.Visible = True
            sprToTCR.ZOrder 0
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'fpToTcrCD.SetFocus
    cboToSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
    
End Sub



Private Sub txtToTcrCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyBack
            fpToTcrCD.Text = ""
            cboToSubjCD.Clear
            
            Exit Sub
            
        Case vbKeyEscape
            sprToTCR.Visible = False
            Exit Sub
                
        Case vbKeyReturn
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM "
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtToTcrCD.Text) & "%'"
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtToTcrCD.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtToTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   ORDER BY ACID, TCRCD, SUBJCD"
        
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpToTcrCD.Text = ""
            txtToTcrCD.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpToTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpToTcrCD.Text = Trim(.Fields("TCRCD"))
            txtToTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtToTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpToTcrCD.Text) <> "" Then Call Find_To_TmrChg_Subj(fpToTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprToTCR.Visible = True
            sprToTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprToTCR.MaxRows = sprToTCR.MaxRows + 1
                sprToTCR.Row = sprToTCR.MaxRows
                
                sprToTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprToTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprToTCR.Top = FraTo.Top + fpToTcrCD.Top + fpToTcrCD.Height
            sprToTCR.Left = FraTo.Left + fpToTcrCD.Left
            sprToTCR.Visible = True
            sprToTCR.ZOrder 0
    
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'txtToTcrCD.SetFocus
    cboToSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub

Private Sub txtToTcrCD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, TCRNM"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtToTcrCD.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtToTcrCD.Text) & "%'"
            End If
            sStr = sStr & "   GROUP BY ACID, TCRCD, TCRNM "
            sStr = sStr & "   ORDER BY ACID, TCRCD "
            
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            fpToTcrCD.Text = ""
            txtToTcrCD.Text = ""
        
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            fpToTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpToTcrCD.Text = Trim(.Fields("TCRCD"))
            txtToTcrCD.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtToTcrCD.Text = Trim(.Fields("TCRNM"))
            
            If Trim(fpToTcrCD.Text) <> "" Then Call Find_To_TmrChg_Subj(fpToTcrCD.Text)
            
        ElseIf .RecordCount > 1 Then
            sprToTCR.Visible = True
            sprToTCR.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprToTCR.MaxRows = sprToTCR.MaxRows + 1
                sprToTCR.Row = sprToTCR.MaxRows
                
                sprToTCR.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprToTCR.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprToTCR, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
            
            sprToTCR.Top = FraTo.Top + fpToTcrCD.Top + fpToTcrCD.Height
            sprToTCR.Left = FraTo.Left + fpToTcrCD.Left
            sprToTCR.Visible = True
            sprToTCR.ZOrder 0
    
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'txtToTcrCD.SetFocus
    cboToSubjCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ȸ"
End Sub



Private Sub sprToTCR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            sprToTCR.Visible = False
            
    End Select
End Sub

Private Sub sprToTCR_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprToTCR
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub

Private Sub sprToTCR_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprToTCR
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        '>> ������ �����ֱ�
        .Row = Row
        .Col = 1:       fpToTcrCD.Text = Trim(.Text)
        .Col = 2:       txtToTcrCD.Text = Trim(.Text)
        
        If Trim(fpToTcrCD.Text) <> "" Then Call Find_To_TmrChg_Subj(fpToTcrCD.Text)
        
        .Visible = False
        
        'fptoTcrCD.SetFocus
        cboToSubjCD.SetFocus
        
    End With
End Sub



'## ������ȸ�� �ش簭���� ������ ��� ��ȸ
Private Sub Find_To_TmrChg_Subj(ByVal aTcr As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String

    Dim sSubjCD     As String
    Dim sSubjNM     As String

    On Error GoTo ErrStmt

    sStr = ""
    sStr = sStr & "  SELECT SUBJCD, SUBJNM"
    sStr = sStr & "    FROM SDTCR01TB"
    sStr = sStr & "   WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD = '" & Trim(aTcr) & "'"
    sStr = sStr & "   GROUP BY SUBJCD, SUBJNM "
    sStr = sStr & "   ORDER BY SUBJCD"

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        If .RecordCount = 0 Then
            
            cboToSubjCD.Clear
            cboToSubjCD.AddItem "����" & Space(30) & "X"
            
        Else
            cboToSubjCD.Clear
            
            .MoveFirst

            For nRec = 1 To .RecordCount Step 1

                sSubjCD = ""
                sSubjNM = ""

                If IsNull(.Fields("SUBJCD")) = False Then sSubjCD = Trim(.Fields("SUBJCD"))
                If IsNull(.Fields("SUBJNM")) = False Then sSubjNM = Trim(.Fields("SUBJNM"))

                cboToSubjCD.AddItem sSubjNM & Space(30) & sSubjCD

                .MoveNext
            Next nRec
        End If
    End With

    If cboToSubjCD.ListCount > 0 Then cboToSubjCD.ListIndex = 0

    Set DBCmd = Nothing
    Set DBRec = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    On Error GoTo 0
    MsgBox "������ ���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���� ������ȸ"
End Sub



'## ����, ����, ����, ���ÿ� �ش��ϴ� ���� ��ȸ
Private Sub fpToBan_GotFocus()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String
    Dim sWeek       As String

    On Error GoTo ErrStmt

    If Trim(fpToTcrCD.UnFmtText) = "" Then Exit Sub
    If Trim(Right(cboToSubjCD.Text, 30)) = "" Or _
       Trim(Right(cboToSubjCD.Text, 30)) = "X" Then Exit Sub
    If Trim(txtToWeek.Text) = "" Then Exit Sub
    If fpToLesson.Value < 1 Or fpToLesson.Value > 10 Then Exit Sub

    Select Case Trim(txtToWeek.Text)
        Case "1"
            sWeek = "2"
        Case "2"
            sWeek = "3"
        Case "3"
            sWeek = "4"
        Case "4"
            sWeek = "5"
        Case "5"
            sWeek = "6"
        Case "6"
            sWeek = "7"
        Case "7"
            sWeek = "1"
    End Select

    sStr = ""
    sStr = sStr & "  SELECT A.LSNCD, "
    
    Select Case Trim(basModule.SchCD)
        Case "N"
            sStr = sStr & " SUBSTR(B.KAEYOL,2,1)||LSNCDNM AS BAN"
        Case "S"
            sStr = sStr & " SUBSTR(B.KAEYOL,2,1)||LSNCDNM AS BAN"
        Case "K"
            sStr = sStr & " SUBSTR(GET_SUBJNM(A.ACID, A.TCRCD, A.SUBJCD), 1, 1)||B.LSNCDNM AS BAN "
    End Select
    
    sStr = sStr & "    FROM SDTRX50TB A, "
    
    sStr = sStr & "         (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 �߰�
    sStr = sStr & "            FROM SDLSN01TB "
    sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "          UNION"
    sStr = sStr & "          SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
    sStr = sStr & "            FROM SDLSN02TB "
    sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "         ) B "
    
    sStr = sStr & "   Where A.ACID  = B.ACID"
    sStr = sStr & "     AND A.LSNCD = B.LSNCD"
    sStr = sStr & "     AND A.YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "     AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TCRCD   = '" & Trim(fpToTcrCD.UnFmtText) & "'"
    sStr = sStr & "     AND SUBJCD  = '" & Trim(Right(cboToSubjCD.Text, 30)) & "'"
    sStr = sStr & "     AND WEEKS   = " & sWeek
    sStr = sStr & "     AND LESSON  = " & Trim(CStr(fpToLesson.UnFmtText))

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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec
        If .RecordCount = 0 Then
            txtToLsnCD.Text = ""
            fpToBan.Text = ""
            
        ElseIf .RecordCount = 1 Then
            .MoveFirst
            
            txtToLsnCD.Text = ""
                If IsNull(.Fields("LSNCD")) = False Then
                    txtToLsnCD.Text = Trim(.Fields("LSNCD"))
                    txtToLsnCD.Text = Trim(.Fields("LSNCD"))        '< �⺻��
                End If
            fpToBan.Text = ""
                If IsNull(.Fields("BAN")) = False Then
                    fpToBan.Text = Trim(.Fields("BAN"))
                    fpToBan.Text = Trim(.Fields("BAN"))             '< �⺻��
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
    MsgBox "������ �ð�ǥ ��ϳ��� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ����"


End Sub






'######################################################################################################################
' �ð�ó��
'######################################################################################################################
Private Sub sprTmr_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sLsnCDNM    As String
    Dim sWeek       As String
    Dim sLesson     As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprSEL
        .MaxRows = 0
    End With
    
    With sprTmr
        .Enabled = False
        
        sLsnCDNM = Trim(fpFromBan.UnFmtText)
        
        Select Case Col
            Case 1
                sWeek = "2"
            Case 2
                sWeek = "3"
            Case 3
                sWeek = "4"
            Case 4
                sWeek = "5"
            Case 5
                sWeek = "6"
            Case 6
                sWeek = "7"
            Case 7
                sWeek = "1"
        End Select
        sLesson = Trim(CStr(Row))
        
        Call Select_tmr_Data(sLsnCDNM, sWeek, sLesson)
        
        .Enabled = True
    End With
    
End Sub

'## �ð�ǥ ���ó��� ó��
    Private Sub Select_tmr_Data(ByVal aLsnCDNM As String, ByVal aWeek As String, ByVal aLesson As String)
    
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
        
        Dim sSubjCD     As String
        Dim sSubjNM     As String
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT TCRCD , SUBJCD, TCRNM, SUBJNM, LSNCD, LSNCDNM, "
        sStr = sStr & "               WEEKS , DECODE(WEEKS, '2','��','3','ȭ','4','��','5','��','6','��','7','��','1','��') AS WEEKNM, "
        sStr = sStr & "               LESSON "
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       SDLSN01TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
                
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','�ι���','2','�ڿ���','��Ÿ') AS KAEYOLNM,"
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
        sStr = sStr & "         WHERE WEEKS   = '" & aWeek & "'"
        sStr = sStr & "           AND LESSON  = '" & aLesson & "'"
        sStr = sStr & "           AND LSNCDNM = '" & aLsnCDNM & "'"
        sStr = sStr & "        UNION "
        sStr = sStr & "        SELECT TCRCD , SUBJCD, TCRNM, SUBJNM, LSNCD, LSNCDNM, "
        sStr = sStr & "               WEEKS , DECODE(WEEKS, '2','��','3','ȭ','4','��','5','��','6','��','7','��','1','��') AS WEEKNM, "
        sStr = sStr & "               LESSON "
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
                
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       SDLSN01TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
                
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','�ι���','2','�ڿ���','��Ÿ') AS KAEYOLNM,"
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
        sStr = sStr & "         WHERE WEEKS   =  '" & aWeek & "'"
        sStr = sStr & "           AND LESSON  =  '" & aLesson & "'"
        sStr = sStr & "           AND LSNCDNM = '" & aLsnCDNM & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30


        
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                
                sprSEL.MaxRows = .RecordCount
                
                .MoveFirst
                
                If .RecordCount = 1 Then
                    sTmp = " ":       If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                        fpToTcrCD.Text = sTmp
                        
                    sTmp = " ":       If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        fpToBan.Text = sTmp
                    sTmp = " ":       If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        txtToLsnCD.Text = sTmp
                    
                    sTmp = " ":       If IsNull(.Fields("WEEKS")) = False Then sTmp = Trim(.Fields("WEEKS"))
                        Select Case sTmp
                            Case "2"
                                txtToWeek.Text = "1"
                            Case "3"
                                txtToWeek.Text = "2"
                            Case "4"
                                txtToWeek.Text = "3"
                            Case "5"
                                txtToWeek.Text = "4"
                            Case "6"
                                txtToWeek.Text = "5"
                            Case "7"
                                txtToWeek.Text = "6"
                            Case "1"
                                txtToWeek.Text = "7"
                        End Select
                    sTmp = " ":       If IsNull(.Fields("LESSON")) = False Then sTmp = Trim(.Fields("LESSON"))
                        If IsNumeric(sTmp) = True Then fpToLesson.Value = CLng(sTmp)
                    sTmp = " ":       If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        txtToTcrCD.Text = sTmp
                    sTmp = " ":       If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        sSubjNM = sTmp
                    sTmp = " ":       If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        'SKIP
                    sTmp = " ":       If IsNull(.Fields("WEEKNM")) = False Then sTmp = Trim(.Fields("WEEKNM"))
                        'SKIP
                    sTmp = " ":       If IsNull(.Fields("LESSON")) = False Then sTmp = Trim(.Fields("LESSON"))
                        'SKIP
                       
                    sTmp = " ":       If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))          '< ���ڵ�
                        sSubjCD = sTmp
                        
                    cboToSubjCD.Clear
                    cboToSubjCD.AddItem sSubjNM & Space(30) & sSubjCD
                    cboToSubjCD.ListIndex = 0
                    
                End If
                
                For nRec = 1 To .RecordCount Step 1
                    sprSEL.Row = nRec
                    
                    sprSEL.Col = 1
                        sTmp = " ":       If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("WEEKS")) = False Then sTmp = Trim(.Fields("WEEKS"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("LESSON")) = False Then sTmp = Trim(.Fields("LESSON"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                        
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("WEEKNM")) = False Then sTmp = Trim(.Fields("WEEKNM"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                    sprSEL.Col = sprSEL.Col + 1
                        sTmp = " ":       If IsNull(.Fields("LESSON")) = False Then sTmp = Trim(.Fields("LESSON"))
                            Call basFunction.Set_SprType_Text(sprSEL, "CENTER", "LEFT", 60, sTmp)
                        
                    .MoveNext
                Next nRec
            End If
        End With
        
        With sprSEL
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBRec = Nothing
        Set DBParam = Nothing
        
        On Error GoTo 0
    End Sub
    
    
    Private Sub sprSEL_Click(ByVal Col As Long, ByVal Row As Long)
        
        Dim sSubjCD     As String
        Dim sSubjNM     As String
        
        Dim sTmp        As String
        
        If Row < 1 Then Exit Sub
        If Col < 1 Then Exit Sub
        
        
        With sprSEL
            If .MaxRows < 1 Then Exit Sub
            
            If .Tag = "" Then .Tag = "1"
    
            .Row = CLng(.Tag):  .Row2 = .Row
            .Col = 1:           .Col2 = 8
            .BlockMode = True
                .BackColor = &HFFFFFF
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        
            .Row = Row:     .Row2 = .Row
            .Col = 1:       .Col2 = .MaxCols
            .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        
            .Col = 1:               sTmp = Trim(.Text):         fpToTcrCD.Text = sTmp
            .Col = .Col + 1:        sTmp = Trim(.Text):         sSubjCD = sTmp
            .Col = .Col + 1:        sTmp = Trim(.Text):         fpToBan.Text = sTmp
            .Col = .Col + 1:        sTmp = Trim(.Text):         txtToLsnCD.Text = sTmp
            .Col = .Col + 1:        sTmp = Trim(.Text)
                Select Case sTmp
                    Case "2"
                        txtToWeek.Text = "1"
                    Case "3"
                        txtToWeek.Text = "2"
                    Case "4"
                        txtToWeek.Text = "3"
                    Case "5"
                        txtToWeek.Text = "4"
                    Case "6"
                        txtToWeek.Text = "5"
                    Case "7"
                        txtToWeek.Text = "6"
                    Case "1"
                        txtToWeek.Text = "7"
                End Select
            .Col = .Col + 1:        sTmp = Trim(.Text):         fpToLesson.Value = CLng(sTmp)
            
            
            .Col = .Col + 1:        sTmp = Trim(.Text):         txtToTcrCD.Text = sTmp
            .Col = .Col + 1:        sTmp = Trim(.Text):         sSubjNM = sTmp
            
            '������ SKIP
                    
                
            cboToSubjCD.Clear
            cboToSubjCD.AddItem sSubjNM & Space(30) & sSubjCD
            cboToSubjCD.ListIndex = 0
            
        End With
        
    End Sub

    

'######################################################################################################################
' �ð���ȸ
'######################################################################################################################
Private Sub cmdTmr_Click()
     
    Dim sF_TcrCD    As String
    Dim sF_SubjCD   As String
    Dim sF_LsnCD    As String
    Dim sF_LsnCDNM  As String
    
    Dim sF_Weeks    As String
    Dim sF_Lesson   As String
    
    Dim sLsnCD      As String
    Dim sKaeyol     As String
    Dim sLsn        As String
    Dim sGwamok     As String
    
    Dim nRowj       As Long
    Dim nRow        As Long
    Dim nCol        As Long
    
   
    '>> �ð�ǥ �ڵ� üũ
        If Trim(fpYM.UnFmtText) = "" Then
            MsgBox "�ð�ǥ �ڵ带 Ȯ���ϼ���.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
    
    '>> 1. �ݳ��� -> ���ڵ�� �ٲٱ� ( ������ �ð�ǥ ���� )
        If Trim(fpFromBan.UnFmtText) = "" Or Len(fpFromBan.UnFmtText) <> 3 Then
            MsgBox "������ �ð�ǥ �������� �� ������ ��������.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        End If
        
        sKaeyol = "0" & Left(Trim(fpFromBan.UnFmtText), 1)
        sLsn = Right(Trim(fpFromBan.UnFmtText), 2)
        
        Call Get_LsnCD_Data(sLsnCD, sKaeyol, sLsn)
        
        If Trim(sLsnCD) = "" Then
            MsgBox "������ �ð�ǥ �������� �� ������ �ش��ϴ� ������ ������ Ȯ���Ͻʽÿ�.", vbExclamation + vbOKOnly, "�ð�ǥ ����"
            Exit Sub
        Else
            txtFromLsnCD.Text = Trim(sLsnCD)
            sF_LsnCD = txtFromLsnCD.Text
        End If
    
    ' ����üũ
        If Trim(fpFromTcrCD.UnFmtText) = "" Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Len(fpFromTcrCD.UnFmtText) <> 3 Then
            MsgBox "���縦 Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
        If Trim(Right(cboFromSubjCD.Text, 30)) = "X" Then
            MsgBox "������ �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub          '< �������.
        End If
        
        Select Case Trim(txtFromWeek.Text)
            Case "1"
                sF_Weeks = "2"
            Case "2"
                sF_Weeks = "3"
            Case "3"
                sF_Weeks = "4"
            Case "4"
                sF_Weeks = "5"
            Case "5"
                sF_Weeks = "6"
            Case "6"
                sF_Weeks = "7"
            Case "7"
                sF_Weeks = "1"
            Case Else
                MsgBox "������ Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        
        Select Case CLng(fpFromLesson.Text)
            Case 1 To 10
                sF_Lesson = Trim(fpFromLesson.Text)
            Case Else
                MsgBox "���ø� Ȯ���ϼ���", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
                Exit Sub
        End Select
        If Trim(txtFromLsnCD.Text) = "" Then
            MsgBox "������ ������ ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ð�ǥ �����ϱ�"
            Exit Sub
        End If
     
    
    '## ������ ��ȸ
    
    With sprTmr
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
                
                .Row2 = .Row
                .Col2 = .Col
                
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            Next nCol
        Next nRow
    End With
    
    cmdTmr.Tag = Trim(fpFromBan.UnFmtText)
    
    sF_TcrCD = Trim(fpFromTcrCD.UnFmtText)
    sF_SubjCD = Trim(Right(cboFromSubjCD.Text, 30))
    'sF_LsnCD
    'sF_Weeks
    'sF_Lesson
    sF_LsnCDNM = Trim(fpFromBan.UnFmtText)
    
    sGwamok = Get_Gwamok_GBN(sF_TcrCD, sF_SubjCD)       '>> 1. ���� ����
    
    Select Case sGwamok
        Case "10", "20", "30"               '# �� �� ��
            With sprTmr
                For nRowj = 1 To .MaxRows Step 1
                    For nCol = 1 To .MaxCols Step 1
                        .Row = nRowj
                        .Col = nCol
                            .Text = ""
                    Next nCol
                Next nRowj
            End With

            Call Data_TCR(sF_TcrCD, sF_SubjCD)                      '>> 2. ������ ���� VIEW
            Call Data_Lsn(sF_TcrCD, sF_LsnCDNM)                     '>> 3. LSNCDNM �� ���� ������ �ҷ�����.
            Call Data_Teaching(sF_TcrCD, sF_SubjCD)                 '>> 4. ���ǰ��� �ü� �κ�
            Call Data_not_Teaching(sF_TcrCD, sF_SubjCD)             '>> 5. ���ǺҰ��� �ü� �κ�


        Case "40", "50"                     '# �� ��Ž
            With sprTmr
                For nRowj = 1 To .MaxRows Step 1
                    For nCol = 1 To .MaxCols Step 1
                        .Row = nRowj
                        .Col = nCol
                            .Text = ""         '<< 1. �� �������� X �� �����Ұ�
                    Next nCol
                Next nRowj
            End With

            Call Data_TCR(sF_TcrCD, sF_SubjCD)                      '>> 2. ������ ���� VIEW
            Call Data_Lsn(sF_TcrCD, sF_LsnCDNM)                     '>> 3. LSNCDNM �� ���� ������ �ҷ�����.
            Call Data_Teaching_Tamgu("1", sF_LsnCD)                 '>> 4. ���ǰ��� �ü� �κ�
            Call Data_not_Teaching(sF_TcrCD, sF_SubjCD)             '>> 5. ���ǺҰ��� �ü� �κ�
            
    End Select
    
End Sub


'## 4. ������ �ð�ǥ ���� ��ȸ
    Private Sub Data_Teaching_Tamgu(ByVal aAlloc As String, ByVal aLsnCD As String)
        
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
        
        sStr = sStr & "               (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "                     '2009.01.12 �߰�
        sStr = sStr & "                  FROM SDLSN01TB "
        sStr = sStr & "                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION"
        sStr = sStr & "                SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL, DAMIM, BASE_CLASS "
        sStr = sStr & "                  FROM SDLSN02TB "
        sStr = sStr & "                 WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "               ) B "
        
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
    
     
     
    
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        If DBRec.RecordCount < 1 Then
            '## ������ �ð�ǥ ��� ���� ����.
            
        Else
            DBRec.MoveFirst
               
            For nRec = 1 To DBRec.RecordCount Step 1
            
                Set DBRecj = New ADODB.Recordset
            
                sStr = ""
                Select Case aAlloc
                    Case "X"            '< �ð����� �Ұ��� �κ� ����
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
                


                
                DBRecj.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
                Do While DBRecj.State And adStateExecuting
                    DoEvents
                Loop
                
                If DBRecj.RecordCount < 1 Then
                    'NOTHING
                Else
                    DBRecj.MoveFirst
                    For nRecj = 1 To DBRecj.RecordCount Step 1
                        Select Case Trim(DBRecj.Fields("WEEKS"))    '< ����
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
                        nLesson = CLng(DBRecj.Fields("LESSON"))     '< ����
                        
                        sprTmr.Row = nLesson
                        sprTmr.Col = nWeek
                        
                        sTmp = Trim(sprTmr.Text)
                        If sTmp = "" Then
                            sTmp = "O"
                        Else
                            sTmp = "O" & vbCrLf & sTmp
                        End If
                        Call basFunction.Set_SprType_Text(sprTmr, "TOP", "LEFT", 100, sTmp)
                        sprTmr.TypeEditMultiLine = True
                        
                        sprTmr.Row2 = sprTmr.Row
                        sprTmr.Col2 = sprTmr.Col
                        sprTmr.BlockMode = True
                            sprTmr.BackColor = &HFF8080
                            sprTmr.BackColorStyle = BackColorStyleUnderGrid
                        sprTmr.BlockMode = False
                        
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

'## 5. ���ǺҰ��� �ü� �κ�
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
        
   
   
        
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< ����
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
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< ����
                        
                        sprTmr.Row = nLesson
                        sprTmr.Col = nWeek
                        
                        sTmp = Trim(sprTmr.Text)
                        If sTmp = "" Then
                            sTmp = "#"
                        Else
                            sTmp = "#" & vbCrLf & sTmp
                        End If
                        Call basFunction.Set_SprType_Text(sprTmr, "TOP", "LEFT", 100, sTmp)
                        sprTmr.TypeEditMultiLine = True
                        
                        sprTmr.Row2 = sprTmr.Row
                        sprTmr.Col2 = sprTmr.Col
                        sprTmr.BlockMode = True
                            sprTmr.BackColor = TMR051.lblNotTeaching.BackColor
                            sprTmr.BackColorStyle = BackColorStyleUnderGrid
                        sprTmr.BlockMode = False
                        
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
  
  
'## 4. ���ǰ��� �ü� �κ�
    Private Sub Data_Teaching(ByVal aTcrCD As String, ByVal aSubjCD As String)
    
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


        
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< ����
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
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< ����
                        
                        sprTmr.Row = nLesson
                        sprTmr.Col = nWeek
                        
                        sTmp = Trim(sprTmr.Text)
                        If sTmp = "" Then
                            sTmp = "O"
                        Else
                            sTmp = "O" & vbCrLf & sTmp
                        End If
                        Call basFunction.Set_SprType_Text(sprTmr, "TOP", "LEFT", 100, sTmp)
                        sprTmr.TypeEditMultiLine = True
                        
                        sprTmr.Row2 = sprTmr.Row
                        sprTmr.Col2 = sprTmr.Col
                        sprTmr.BlockMode = True
                            sprTmr.BackColor = &HFF8080
                            sprTmr.BackColorStyle = BackColorStyleUnderGrid
                        sprTmr.BlockMode = False
                        
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
  
'## 3. ������ ���� ����
    Private Sub Data_Lsn(ByVal aTcrCD As String, ByVal aLsnCDNM As String)
    
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        Dim sStr        As String
        
        Dim sTmp        As String
        Dim nTmp        As Long
        
        Dim ni          As Long
        Dim nRec        As Long
        
        Dim sTcrCD      As String
        Dim sSubjCD     As String
        Dim sLsnCDNM    As String
        
        Dim nWeek       As Long
        Dim nLesson     As Long
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT TCRNM, SUBJNM, LSNCDNM, WEEKS, LESSON"
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
                
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       SDLSN01TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','�ι���','2','�ڿ���','��Ÿ') AS KAEYOLNM,"
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
        sStr = sStr & "         WHERE LSNCDNM =  '" & aLsnCDNM & "'"
        sStr = sStr & "           AND TCRCD   <> '" & aTcrCD & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
  
  
        
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< ����
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
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< ����
                        
                        
                        sTcrCD = " ":       If IsNull(.Fields("TCRNM")) = False Then sTcrCD = Trim(.Fields("TCRNM"))
                        sSubjCD = " ":      If IsNull(.Fields("SUBJNM")) = False Then sSubjCD = Trim(.Fields("SUBJNM"))
                        sLsnCDNM = " ":     If IsNull(.Fields("LSNCDNM")) = False Then sLsnCDNM = Trim(.Fields("LSNCDNM"))
                        
                        sprTmr.Row = nLesson
                        sprTmr.Col = nWeek
                            
                            sTmp = Trim(sprTmr.Text)
                            If sTmp = "" Then
                                sTmp = sLsnCDNM & ", " & sSubjCD & ", " & sTcrCD
                            Else
                                sTmp = sTmp & vbCrLf & sLsnCDNM & ", " & sSubjCD & ", " & sTcrCD
                            End If
                            
                            Call basFunction.Set_SprType_Text(sprTmr, "TOP", "LEFT", 100, sTmp)
                            sprTmr.TypeEditMultiLine = True
                            
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


'## 2. ������ ���� ����
    Private Sub Data_TCR(ByVal aTcrCD As String, ByVal aSubjCD As String)
    
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        Dim sStr        As String
        
        Dim sTmp        As String
        Dim nTmp        As Long
        
        Dim ni          As Long
        Dim nRec        As Long
        
        Dim sTcrCD      As String
        Dim sSubjCD     As String
        Dim sLsnCDNM    As String
        
        Dim nWeek       As Long
        Dim nLesson     As Long
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "        SELECT TCRNM, SUBJNM, LSNCDNM, WEEKS, LESSON"
        sStr = sStr & "          FROM (SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       SDLSN01TB B"
        sStr = sStr & "                 WHERE A.ACID  = B.ACID"
        sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
        sStr = sStr & "                   AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "                UNION ALL"
        sStr = sStr & "                SELECT A.LSNCD, A.LSNNM,"
        sStr = sStr & "                       B.KAEYOL,"
        sStr = sStr & "                       DECODE(B.KAEYOL,'01','�ι���','02','�ڿ���','03','��ü��') AS KAEYOLNM,"
        sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
        sStr = sStr & "                       B.DAMIM,"
        sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
        
        Select Case Trim(basModule.SchCD)
            Case "N"
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
        sStr = sStr & "                       DECODE(SUBSTR(PRT_KAEYOL,1,1),'1','�ι���','2','�ڿ���','��Ÿ') AS KAEYOLNM,"
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
        sStr = sStr & "           AND SUBJCD = '" & aSubjCD & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        


        
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            If .RecordCount > 0 Then
                .MoveFirst
                
                For nRec = 1 To .RecordCount Step 1
                    
                        Select Case Trim(DBRec.Fields("WEEKS"))    '< ����
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
                        nLesson = CLng(DBRec.Fields("LESSON"))     '< ����
                        
                        
                        sTcrCD = " ":       If IsNull(.Fields("TCRNM")) = False Then sTcrCD = Trim(.Fields("TCRNM"))
                        sSubjCD = " ":      If IsNull(.Fields("SUBJNM")) = False Then sSubjCD = Trim(.Fields("SUBJNM"))
                        sLsnCDNM = " ":     If IsNull(.Fields("LSNCDNM")) = False Then sLsnCDNM = Trim(.Fields("LSNCDNM"))
                        
                        sprTmr.Row = nLesson
                        sprTmr.Col = nWeek
                            
                            sTmp = Trim(sprTmr.Text)
                            If sTmp = "" Then
                                sTmp = sLsnCDNM & ", " & sSubjCD & ", " & sTcrCD
                            Else
                                sTmp = sTmp & vbCrLf & sLsnCDNM & ", " & sSubjCD & ", " & sTcrCD
                            End If
                            
                            Call basFunction.Set_SprType_Text(sprTmr, "TOP", "LEFT", 100, sTmp)
                            sprTmr.TypeEditMultiLine = True
                            
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

'## 1. ���񱸺� ��ȸ
    Private Function Get_Gwamok_GBN(ByVal aTcrCD As String, ByVal aSubjCD As String) As String
        Dim DBCmd       As ADODB.Command
        Dim DBRec       As ADODB.Recordset
        Dim DBParam     As ADODB.Parameter
        
        Dim nLength     As Long
        
        Dim sStr        As String
        Dim sRet        As String
        
        Dim ni          As Integer
        
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & " SELECT MAX(SUBJGBN) AS SUBJGBN"
        sStr = sStr & "   From SDTCR01TB"
        sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "    AND TCRCD  = '" & Trim(aTcrCD) & "'"
        sStr = sStr & "    AND SUBJCD = '" & Trim(aSubjCD) & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        


        
        '>> �п�
    '    sTmp = Trim(basModule.SchCD)
    '    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '        Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        sRet = ""
        
        With DBRec
            If .RecordCount = 1 Then
                .MoveFirst
                
                sRet = " ":     If IsNull(.Fields("SUBJGBN")) = False Then sRet = Trim(.Fields("SUBJGBN"))
            End If
        End With
        
ErrStmt:
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        On Error GoTo 0
        
        Get_Gwamok_GBN = sRet
        
    End Function










'## B ##

'
''## ���� ���� ���� ã�Ƴ�
'Private Sub Find_Old_Tmr_Data(ByRef aTcrCD As String, ByRef aSubjCD As String, _
'                              ByVal aLsnCD As String, ByVal aWeeks As String, ByVal aLesson As String)
'
'    Dim DBCmd       As ADODB.Command
'    Dim DBRec       As ADODB.Recordset
'    Dim DBParam     As ADODB.Parameter
'
'    Dim nLength     As Long
'    Dim sStr        As String
'    Dim ni          As Integer
'    Dim nRec        As Long
'
'    Dim sTmp        As String
'
'    On Error GoTo ErrStmt
'
'    aTcrCD = ""
'    aSubjCD = ""
'
'    sStr = ""
'    sStr = sStr & "  SELECT TCRCD, SUBJCD "
'    sStr = sStr & "    FROM SDTRX50TB"
'    sStr = sStr & "   Where YM     = '" & Trim(fpYM.UnFmtText) & "'"
'    sStr = sStr & "     AND ACID   = '" & Trim(basModule.SchCD) & "'"
'    sStr = sStr & "     AND WEEKS  = " & aWeeks
'    sStr = sStr & "     AND LESSON = " & aLesson
'    sStr = sStr & "     AND LSNCD  = " & aLsnCD
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
''    ' ACID
''        sTmp = Trim(basModule.SchCD)
''        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
'    Do While DBRec.State And adStateExecuting
'        DoEvents
'    Loop
'
'    With DBRec
'        If .RecordCount = 1 Then
'            .MoveFirst
'
'            aTcrCD = "":    If IsNull(.Fields("TCRCD")) = False Then aTcrCD = Trim(.Fields("TCRCD"))
'            aSubjCD = "":   If IsNull(.Fields("SUBJCD")) = False Then aSubjCD = Trim(.Fields("SUBJCD"))
'        End If
'    End With
'
'ErrStmt:
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'End Sub
'












'## A ##

'   ## �뷮�� ��û����
'    '** ����Ǿ��� �׸��� ���� ���� ( A -> B ���� B ) **
'        If Trim(sT_TcrCD) <> "" And Trim(sT_SubjCD) <> "" Then
'            bRet = Del_TMR_Data(sT_TcrCD, sT_SubjCD, sT_LsnCD, sT_Weeks, sT_Lesson)
'
'            If bRet = True Then
'                ' ����,���� & �� ���� ����
'                For nRow = 1 To TMR051.sprTmr_Lsn.MaxRows Step 1
'                    TMR051.sprTmr_Lsn.Row = nRow
'                    TMR051.sprTmr_Lsn.Col = SpreadHeader + 1        '< ����
'
'                    If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_Weeks, vbTextCompare) = 0 Then
'                        nr_Chk = TMR051.sprTmr_Lsn.Row              '< row ��
'
'                        TMR051.sprTmr_Lsn.Col = SpreadHeader + 2        '< lesson
'
'                        If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_Lesson, vbTextCompare) = 0 Then
'
'                            For nCol = 1 To TMR051.sprTmr_Lsn.MaxCols Step 1
'                                TMR051.sprTmr_Lsn.Col = nCol
'                                TMR051.sprTmr_Lsn.Row = SpreadHeader + 1
'
'                                If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_LsnCD, vbTextCompare) = 0 Then
'                                    nc_Chk = TMR051.sprTmr_Lsn.Col
'
'                                    TMR051.sprTmr_Lsn.Row = nr_Chk
'                                    TMR051.sprTmr_Lsn.Col = nc_Chk
'                                        TMR051.sprTmr_Lsn.Text = ""
'
'                                    Exit For
'                                End If
'                            Next nCol
'                        End If
'                    End If
'                Next nRow
'
'                ' ���� & ���� ���� ����
'                For nRow = 1 To TMR051.sprTmr_Tcr.MaxRows Step 1
'                    TMR051.sprTmr_Tcr.Row = nRow
'                    TMR051.sprTmr_Tcr.Col = SpreadHeader
'
'                    If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_TcrCD, vbTextCompare) = 0 Then
'                        TMR051.sprTmr_Tcr.Col = SpreadHeader + 1
'
'                        If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_SubjCD, vbTextCompare) = 0 Then
'                            nr_Chk = TMR051.sprTmr_Tcr.Row
'
'                            For nCol = 1 To TMR051.sprTmr_Tcr.MaxCols Step 1
'                                TMR051.sprTmr_Tcr.Col = nCol
'                                TMR051.sprTmr_Tcr.Row = SpreadHeader + 1
'
'                                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_Weeks, vbTextCompare) = 0 Then
'                                    TMR051.sprTmr_Tcr.Row = SpreadHeader + 2
'
'                                    If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_Lesson, vbTextCompare) = 0 Then
'                                        nc_Chk = TMR051.sprTmr_Tcr.Col
'
'                                        TMR051.sprTmr_Tcr.Row = nr_Chk
'                                        TMR051.sprTmr_Tcr.Col = nc_Chk
'                                            TMR051.sprTmr_Tcr.Text = ""
'
'                                        Exit For
'                                    End If
'                                End If
'                            Next nCol
'                        End If
'                    End If
'                Next nRow
'
'            End If
'        End If















' ' ����,���� & �� ���� ���
'            For nRow = 1 To TMR051.sprTmr_Lsn.MaxRows Step 1
'                TMR051.sprTmr_Lsn.Row = nRow
'                TMR051.sprTmr_Lsn.Col = SpreadHeader + 1        '< ����
'
'                If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_Weeks, vbTextCompare) = 0 Then
'                    nr_Chk = TMR051.sprTmr_Lsn.Row              '< row ��
'
'                    TMR051.sprTmr_Lsn.Col = SpreadHeader + 2        '< lesson
'
'                    If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_Lesson, vbTextCompare) = 0 Then
'
'                        For nCol = 1 To TMR051.sprTmr_Lsn.MaxCols Step 1
'                            TMR051.sprTmr_Lsn.Col = nCol
'                            TMR051.sprTmr_Lsn.Row = SpreadHeader + 1
'
'                            If StrComp(Trim(TMR051.sprTmr_Lsn.Text), sT_LsnCD, vbTextCompare) = 0 Then
'                                nc_Chk = TMR051.sprTmr_Lsn.Col
'
'                                TMR051.sprTmr_Lsn.Row = nr_Chk
'                                TMR051.sprTmr_Lsn.Col = nc_Chk
'                                    TMR051.sprTmr_Lsn.Text = sF_SubjNM & "," & sF_TcrNM
'
'                                Exit For
'                            End If
'                        Next nCol
'                    End If
'                End If
'            Next nRow
'
'            ' ���� & ���� ���� ���
'            For nRow = 1 To TMR051.sprTmr_Tcr.MaxRows Step 1
'                TMR051.sprTmr_Tcr.Row = nRow
'                TMR051.sprTmr_Tcr.Col = SpreadHeader
'
'                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_TcrCD, vbTextCompare) = 0 Then
'                    TMR051.sprTmr_Tcr.Col = SpreadHeader + 1
'
'                    If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sF_SubjCD, vbTextCompare) = 0 Then
'                        nr_Chk = TMR051.sprTmr_Tcr.Row
'
'                        For nCol = 1 To TMR051.sprTmr_Tcr.MaxCols Step 1
'                            TMR051.sprTmr_Tcr.Col = nCol
'                            TMR051.sprTmr_Tcr.Row = SpreadHeader + 1
'
'                            If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_Weeks, vbTextCompare) = 0 Then
'                                TMR051.sprTmr_Tcr.Row = SpreadHeader + 2
'
'                                If StrComp(Trim(TMR051.sprTmr_Tcr.Text), sT_Lesson, vbTextCompare) = 0 Then
'                                    nc_Chk = TMR051.sprTmr_Tcr.Col
'
'                                    TMR051.sprTmr_Tcr.Row = nr_Chk
'                                    TMR051.sprTmr_Tcr.Col = nc_Chk
'                                        TMR051.sprTmr_Tcr.Text = sT_LsnCDNM
'
'                                    Exit For
'                                End If
'                            End If
'                        Next nCol
'                    End If
'                End If
'            Next nRow
