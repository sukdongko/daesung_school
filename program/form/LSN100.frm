VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form LSN100 
   Caption         =   "�ð�ǥ ����� >> �� �����ϱ�"
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
      BorderStyle     =   0  '����
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
         BackStyle       =   0  '����
         Caption         =   "�� ó�� �ӽ� Spread"
         Height          =   210
         Left            =   180
         TabIndex        =   42
         Top             =   60
         Width           =   2265
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame7"
      Height          =   3645
      Left            =   30
      TabIndex        =   37
      Top             =   6000
      Width           =   15435
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame6"
         Height          =   3585
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   15375
         Begin VB.CommandButton cmdOrdGwamok_View 
            Caption         =   "�л���û���� ��ģ���� ����"
            BeginProperty Font 
               Name            =   "����ü"
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
            Caption         =   "�� �����ϱ� (�ݼ��� ��ȸ�ϱ� Ŭ�� -> �� ó���� �л��� ���� -> �����ϱ� Ŭ���ϼ���.)"
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
            Left            =   3000
            TabIndex        =   25
            Top             =   30
            Width           =   8505
         End
         Begin VB.CommandButton cmdDeleteClass 
            Caption         =   "���� �� ��ϳ��� �����ϱ�"
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
            Caption         =   "�� ������ȸ"
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
            BackStyle       =   0  '����
            Caption         =   "�������� <delete Ű ����>"
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   60
            Top             =   3300
            Width           =   3405
         End
         Begin VB.Label Label6 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
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
      BorderStyle     =   0  '����
      Caption         =   "Frame5"
      Height          =   615
      Left            =   30
      TabIndex        =   34
      Top             =   30
      Width           =   15465
      Begin VB.Frame Frame4 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame4"
         Height          =   555
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   15405
         Begin VB.CommandButton cmdinput_Class 
            Caption         =   "�� ����ϱ�"
            Height          =   495
            Left            =   13650
            TabIndex        =   43
            Top             =   30
            Width           =   1725
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   $"LSN100.frx":4621
            Height          =   375
            Left            =   5160
            TabIndex        =   46
            Top             =   90
            Width           =   8805
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
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
      BorderStyle     =   0  '����
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   30
      TabIndex        =   29
      Top             =   660
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   5235
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   15375
         Begin VB.Frame Frame10 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
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
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   16
               Top             =   120
               Width           =   945
            End
            Begin VB.ComboBox cboGwamok 
               Height          =   300
               Index           =   1
               Left            =   6090
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   15
               Top             =   120
               Width           =   945
            End
            Begin VB.ComboBox cboGwamok 
               Height          =   300
               Index           =   0
               Left            =   5160
               Style           =   2  '��Ӵٿ� ���
               TabIndex        =   14
               Top             =   120
               Width           =   945
            End
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
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�����ȣ"
               Height          =   210
               Index           =   6
               Left            =   1080
               TabIndex        =   70
               Top             =   0
               Width           =   765
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "Ÿ��"
               Height          =   210
               Index           =   5
               Left            =   1890
               TabIndex        =   68
               Top             =   -30
               Width           =   465
            End
            Begin VB.Label Label18 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   7140
               TabIndex        =   63
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label14 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   6180
               TabIndex        =   62
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label13 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   5280
               TabIndex        =   61
               Top             =   0
               Width           =   465
            End
            Begin VB.Label Label19 
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
               TabIndex        =   58
               Top             =   105
               Width           =   645
            End
            Begin VB.Label Label7 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�հ�"
               Height          =   210
               Left            =   4470
               TabIndex        =   57
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "���"
               Height          =   210
               Index           =   1
               Left            =   2580
               TabIndex        =   56
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Index           =   2
               Left            =   3210
               TabIndex        =   55
               Top             =   -15
               Width           =   465
            End
            Begin VB.Label Label12 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�ܱ���"
               Height          =   210
               Index           =   3
               Left            =   3810
               TabIndex        =   54
               Top             =   -15
               Width           =   615
            End
         End
         Begin VB.CommandButton cmdDelStdClass 
            Caption         =   "�����л� �� �����ϱ�"
            Height          =   315
            Left            =   12750
            TabIndex        =   27
            Top             =   4860
            Width           =   2595
         End
         Begin VB.CommandButton cmdNotProcDataSelect 
            Caption         =   "�� ��������"
            Height          =   315
            Left            =   13920
            TabIndex        =   20
            Top             =   495
            Width           =   1245
         End
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "����"
            Height          =   315
            Left            =   14250
            TabIndex        =   21
            Top             =   840
            Width           =   885
         End
         Begin VB.CommandButton cmdFindStd 
            Caption         =   "�л� ��ȸ�ϱ�"
            Height          =   405
            Left            =   210
            TabIndex        =   0
            Top             =   30
            Width           =   1605
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   4050
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   2
            Top             =   105
            Width           =   915
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   2550
            Style           =   2  '��Ӵٿ� ���
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
            Begin VB.Label Label9 
               BackStyle       =   0  '����
               Caption         =   "�հ�             ����"
               Height          =   210
               Left            =   5010
               TabIndex        =   51
               Top             =   180
               Width           =   1995
            End
            Begin VB.Label Label15 
               BackStyle       =   0  '����
               Caption         =   "���            �̻�/"
               Height          =   210
               Left            =   60
               TabIndex        =   50
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label16 
               BackStyle       =   0  '����
               Caption         =   "����            �̻�/"
               Height          =   210
               Left            =   1680
               TabIndex        =   49
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label17 
               BackStyle       =   0  '����
               Caption         =   "�ܱ���            �̻�/"
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00D2EAF5&
            Height          =   465
            Left            =   8130
            TabIndex        =   33
            Top             =   390
            Width           =   4605
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "��ü�л�"
               Height          =   315
               Index           =   2
               Left            =   3360
               TabIndex        =   19
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "�� ������ �л�"
               Height          =   315
               Index           =   1
               Left            =   1740
               TabIndex        =   18
               Top             =   120
               Width           =   1695
            End
            Begin VB.OptionButton optClass 
               BackColor       =   &H00D2EAF5&
               Caption         =   "�� ������ �л�"
               Height          =   315
               Index           =   0
               Left            =   90
               TabIndex        =   17
               Top             =   120
               Width           =   1695
            End
         End
         Begin VB.Label Label12 
            BackStyle       =   0  '����
            Caption         =   "�������� <delete Ű ����>"
            Height          =   210
            Index           =   4
            Left            =   9840
            TabIndex        =   66
            Top             =   4950
            Width           =   3405
         End
         Begin VB.Label Label46 
            BackStyle       =   0  '����
            Caption         =   "��ȸ�ο�"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   7740
            TabIndex        =   65
            Top             =   4950
            Width           =   975
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '����
            Caption         =   $"LSN100.frx":8B9D
            Height          =   360
            Left            =   150
            TabIndex        =   59
            Top             =   4860
            Width           =   11175
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ            ����             ����"
            Height          =   210
            Left            =   5040
            TabIndex        =   52
            Top             =   150
            Width           =   3075
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   3030
            TabIndex        =   32
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��������"
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
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : LSN100
'   �� ��  �� �� : �ð�ǥ ����� >> �� �����ϱ�
'
'   ��   ��   �� : 2007/10/22
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private Type tClass
    CLSCD   As String
    CLSNM   As String
End Type
Private Const nRowHeight = 14


Private Sub cmdOrdGwamok_View_Click()
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "�ݺ� ���� ��û���� ��ȸ�� �Ͻʽÿ�.", vbExclamation + vbOKOnly, "�л���û���� ��ģ���� ����"
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
            .AddItem "��  ü" & Space(30) & "ALL"
            .AddItem "������" & Space(30) & "0"
            .AddItem "������" & Space(30) & "1"
            .ListIndex = 0
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            '.AddItem "��ü" & Space(30) & "03"
            
            .ListIndex = 0
            
            txtKaeyol.Text = Trim(cboKaeyol.Text)
        End With
        
        Call init_Form
        
    Me.Tag = ""
    
End Sub

'## �������
Private Sub cboKaeyol_Click()
    Dim sTmp        As String
    Dim ni          As Integer
    
    txtKaeyol.Text = Trim(cboKaeyol.Text)
    
    With sprSTD
        Select Case Trim(Right(cboKaeyol.Text, 30))
            Case "01", "03"         '<< �ι�
                
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                '.MaxCols = 21
                .MaxCols = 26           '< 2007.12.17
                
                .Col = 1:           .Text = "�л�":         .ColWidth(.Col) = 7.2
                .Col = .Col + 1:    .Text = "�л���":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                
                '< 2007.12.17 ------------------------------------------------------
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "�ܱ�":         .ColWidth(.Col) = 4
                .Col = .Col + 1:    .Text = "�հ�":         .ColWidth(.Col) = 4
                '-------------------------------------------------------------------
                
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "�ѱ�":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "�����":       .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "��ġ":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "�繮":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                
                .Col = .Col + 1:    .Text = "��2��":        .ColWidth(.Col) = 4.5
               
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 4.5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 4.5
                
                .Col = .Col + 1:    .Text = "�ݳֱ�":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "��������":     .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 6
                
            Case "02"       '<< �ڿ�
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                '.MaxCols = 18
                .MaxCols = 23           '< 2007.12.17
                
                .Col = 1:           .Text = "�л�":         .ColWidth(.Col) = 7.2
                .Col = .Col + 1:    .Text = "�л���":       .ColWidth(.Col) = 7
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 6
                
                '< 2007.12.17 ------------------------------------------------------
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�ܱ�":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�հ�":         .ColWidth(.Col) = 5
                '-------------------------------------------------------------------
                
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "�ݳֱ�":       .ColWidth(.Col) = 7.3
                .Col = .Col + 1:    .Text = "��������":     .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5.9
                
        End Select
        
        .MaxRows = 0
    End With
    
    
    With sprClass
        Select Case Trim(Right(cboKaeyol.Text, 30))
            Case "01", "03"         '<< �ι�
                
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 21
                
                .Col = 1:           .Text = "��":           .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�ݸ�":         .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "�ѿ�":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "�����ο�":     .ColWidth(.Col) = 8
                
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�ѱ�":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�����":       .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��ġ":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�繮":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "��2��":        .ColWidth(.Col) = 5
               
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                
            Case "02"       '<< �ڿ�
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 18
                
                .Col = 1:           .Text = "��":           .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "�ݸ�":         .ColWidth(.Col) = 6
                
                .Col = .Col + 1:    .Text = "�ѿ�":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 6
                .Col = .Col + 1:    .Text = "�����ο�":     .ColWidth(.Col) = 8
                
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                
                .Col = .Col + 1:    .Text = "���":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "����":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��Ž":         .ColWidth(.Col) = 5
                
        End Select
        
        .MaxRows = 0
    End With
    
    
    For ni = 0 To 2 Step 1
        With cboGwamok(ni)
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03"         '<< �ι�
                    .AddItem "����" & Space(30) & "X"
                    .AddItem "����" & Space(30) & "8"
                    .AddItem "����" & Space(30) & "9"
                    .AddItem "����" & Space(30) & "10"
                    .AddItem "�ѱ�" & Space(30) & "11"
                    .AddItem "�����" & Space(30) & "12"
                    .AddItem "����" & Space(30) & "13"
                    .AddItem "����" & Space(30) & "14"
                    .AddItem "��ġ" & Space(30) & "15"
                    .AddItem "�繮" & Space(30) & "16"
                    .AddItem "����" & Space(30) & "17"
                    .AddItem "����" & Space(30) & "18"
                                         
'                    .AddItem "��2��" & Space(30) & "19"
'
'                    .AddItem "���" & Space(30) & "20"
'                    .AddItem "����" & Space(30) & "21"
'                    .AddItem "��Ž" & Space(30) & "22"
'                    .AddItem "��Ž" & Space(30) & "23"
                Case "02"
                    .AddItem "����" & Space(30) & "X"
                    .AddItem "��1" & Space(30) & "8"
                    .AddItem "ȭ1" & Space(30) & "9"
                    .AddItem "��1" & Space(30) & "10"
                    .AddItem "��1" & Space(30) & "11"
                    .AddItem "��2" & Space(30) & "12"
                    .AddItem "ȭ2" & Space(30) & "13"
                    .AddItem "��2" & Space(30) & "14"
                    .AddItem "��2" & Space(30) & "15"
                    
                    
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















'>> �л� ��ȸ
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
    
    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '01'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '02'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '03'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "             '04'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         END END SEL4,"
    
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "             '05'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "             '06'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "             '07'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "             '08'"
    sStr = sStr & "         ELSE CASE WHEN A.SEL3 > ' ' AND INSTR(A.SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         END END SEL8,"
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"       '<< �ι�
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
            sStr = sStr & "             '09'"
            sStr = sStr & "         END SEL9,"
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
            sStr = sStr & "             '10'"
            sStr = sStr & "         END SEL10,"
            sStr = sStr & "         CASE WHEN A.SEL1 > ' ' AND INSTR(A.SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
            sStr = sStr & "             '11'"
            sStr = sStr & "         END SEL11,"
    End Select
    
    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
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
    
    sStr = sStr & "      /* ��� */"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'91|') > 0 THEN         /* ��� */"
    sStr = sStr & "             '91'"
    sStr = sStr & "         END SEL_N1,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'92|') > 0 THEN         /* ���� */"
    sStr = sStr & "             '92'"
    sStr = sStr & "         END SEL_N2,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'93|') > 0 THEN         /* ��Ž */"
    sStr = sStr & "             '93'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(A.SEL5,'94|') > 0 THEN         /* ��Ž */"
    sStr = sStr & "             '94'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         GET_LSNNM(A.ACID, A.SEL_CLASS) AS LSNNM, "
    
    '< 2007.12.17 ----------------------------------------------------------------------------------------
    sStr = sStr & "         A.K_NUM, A.M_NUM, A.E_NUM, A.TOT_NUM "
    sStr = sStr & "         , DECODE(B.MU_TYPE,'1','1����','2','6��','3','9��','4','6��','5','9��') AS MU_TYPE "
    '-----------------------------------------------------------------------------------------------------
    
    sStr = sStr & "    FROM CLTTL01TB A, CLSTD01TB B"
    sStr = sStr & "   WHERE A.SCHNO = B.SCHNO "
    
    sStr = sStr & "     AND A.ACID  = B.ACID  "
    
    sStr = sStr & "     AND A.SCHNO > ' ' "
    
    sStr = sStr & "     AND A.ACID = '" & Trim(basModule.SchCD) & "'"
    
'>> �迭
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
'>> ���豸�� (EXMTYPE)
    Select Case Trim(Right(cboExmType.Text, 30))
        Case "ALL"
        
        Case "0"
            sStr = sStr & " AND A.EXMTYPE = '0' "
        Case "1"
            sStr = sStr & " AND A.EXMTYPE = '1' "
    End Select
'>> �����ȣ
    'sStr = sStr & "     AND B.EXMID BETWEEN '" & Format(fpGwanri1.Value, "00000") & "'"
    'sStr = sStr & "                     AND '" & Format(fpGwanri2.Value, "00000") & "'"
'>> �����ȣ            2007.12.17
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND B.EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
    
'>> �հ�
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
        Case "0"        '< ������
        '>> ���
            If fpKor.value > 0 Then
                sStr = sStr & " AND B.K_NUM <= " & Trim(CStr(fpKor.value))
            End If
        '>> ����
            If fpMat.value > 0 Then
                sStr = sStr & " AND B.M_NUM <= " & Trim(CStr(fpMat.value))
            End If
        '>> �ܱ���
            If fpEng.value > 0 Then
                sStr = sStr & " AND B.E_NUM <= " & Trim(CStr(fpEng.value))
            End If
        Case "1"        '< ������
        '>> ���
            If fpKor.value > 0 Then
                sStr = sStr & " AND B.K_NUM >= " & Trim(CStr(fpKor.value))
            End If
        '>> ����
            If fpMat.value > 0 Then
                sStr = sStr & " AND B.M_NUM >= " & Trim(CStr(fpMat.value))
            End If
        '>> �ܱ���
            If fpEng.value > 0 Then
                sStr = sStr & " AND B.E_NUM >= " & Trim(CStr(fpEng.value))
            End If
    End Select
    
'>> �ϷῩ�� : ����Ǹ� YYMM���� ��.
    sStr = sStr & "     AND A.CL_CLOSE IS NULL "
    
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "     AND BIGO1 > 17"                     '< 2009.01.
    Else
        sStr = sStr & "     AND BIGO2 IS NULL"                  '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
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

'   >> ������ȣ
'        sTmp = Format(fpGwanri1.Value, "00000")
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
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
                sprSTD.Col = sprSTD.Col + 1     ' ����
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = CLng(.Fields("K_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' ����
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = CLng(.Fields("M_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' ����
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = CLng(.Fields("E_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                sprSTD.Col = sprSTD.Col + 1     ' �հ�
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = CLng(.Fields("TOT_NUM"))
                        Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 99999, "", nTmp)
                    
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            '-----------------------------------------------------------------------------------------------------------------------------------
            
                
            '>> ���ð��� (��Ž/ ��Ž)
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
                                Case "01":  sTmp = "����"
                                Case "02":  sTmp = "����"
                                Case "03":  sTmp = "����"
                                Case "04":  sTmp = "�ѱ�"
                                Case "05":  sTmp = "�����"
                                Case "06":  sTmp = "����"
                                Case "07":  sTmp = "����"
                                Case "08":  sTmp = "��ġ"
                                Case "09":  sTmp = "�繮"
                                Case "10":  sTmp = "����"
                                Case "11":  sTmp = "����"

                                Case "51":   sTmp = "��1"
                                Case "52":   sTmp = "ȭ1"
                                Case "53":   sTmp = "��1"
                                Case "54":   sTmp = "��1"
                                Case "55":   sTmp = "��2"
                                Case "56":   sTmp = "ȭ2"
                                Case "57":   sTmp = "��2"
                                Case "58":   sTmp = "��2"

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

                            Case "31":  sTmp = "����"
                            Case "32":  sTmp = "�Ͼ�"
                            Case "33":  sTmp = "�����ĳľ�"
                            Case "34":  sTmp = "�Ҿ�"
                            Case "35":  sTmp = "�߱���"
                            Case "36":  sTmp = "�ѹ�"
                            
                            Case "37":  sTmp = "���"
                            Case "38":  sTmp = "����"
                            Case "39":  sTmp = "����"
                            Case "40":  sTmp = "�����"
                            Case "41":  sTmp = "��������"
                            Case "42":  sTmp = "�ƶ���"
                            
                            Case "81":  sTmp = "������"
                            Case "82":  sTmp = "�̻����"
                            Case "83":  sTmp = "Ȯ�����"
                            Case "84":  sTmp = "��������"

                        End Select
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> ���
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
                                Case "91":  sTmp = "���"
                                Case "92":  sTmp = "����"
                                Case "93":  sTmp = "�ܱ���"     '< ����
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
    
    MsgBox "�л� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л���ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ����� �л� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"

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



'>> �� ���õ��� ���� ���� row ����
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








'<< �迭���ý� �迭�� �ش��ϴ� ����ȸ
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
        Case "01"       '<< �ι�
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
                Case "01"       '<< �ι�
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
                Case "01"       '<< �ι�
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
            
            sStr = sStr & "                    /* ��Ž, ��Ž �и� */"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL1,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL2,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL3,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL4,"
            
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL5,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL6,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL7,"
            sStr = sStr & "                        CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END END SEL8,"
            
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01"       '<< �ι�
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL9,"
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL10,"
                    sStr = sStr & "                CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
                    sStr = sStr & "                    1"
                    sStr = sStr & "                ELSE"
                    sStr = sStr & "                    0"
                    sStr = sStr & "                END SEL11,"
            End Select
            
            sStr = sStr & "                 /* ��2�ܱ��� & ���� */"
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
            
            sStr = sStr & "                 /* ��� */"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N1,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N2,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"       '< ����
            sStr = sStr & "                            1"
            sStr = sStr & "                        ELSE"
            sStr = sStr & "                            0"
            sStr = sStr & "                        END SEL_N3,"
            sStr = sStr & "                        CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"             '< ����
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
                Case "01", "03"         '<< �ι�
                    sStr = sStr & "            AND SEL1 > ' '"
                Case "02"               '<< �ڿ�
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
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
                    Case "01"       '<< �ι�
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
    
    MsgBox "�� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л���ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ����� �л� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"

End Sub















'##########################################################################################################
'## �ݼ��� �˰���
'##########################################################################################################

Private Sub cmdProcClass_Click()
    
    Dim nRow            As Long
    Dim nCol            As Long
    
    Dim sClass          As String           ' �ݸ�
    Dim nLimit          As Long             ' �ο���
    Dim nSubj           As Long             ' �����
    
    Dim nTGwamokCnt     As Long
    
    Dim nTotinwon       As Long
    Dim sHeader         As String
    Dim sTmp            As String
    Dim nTmp            As Long
    Dim nMaxStdinwon    As Long
    
    If sprClass.MaxRows = 0 Then
        MsgBox "���� ��ȸ�ϼ���.", vbExclamation + vbOKOnly, "�� ����"
        Exit Sub
    End If
    
    
    
'<<  �����л� -> sprClassDet�� ���� count   >>
    Select Case Trim(Right(txtKaeyol.Text, 30))
        Case "01", "03"         '<< �ι��� : 11 ����
            With sprClassDet
                .MaxCols = 0
                .MaxRows = 0
                
                .MaxCols = 18
                .MaxRows = 2
                
                .Row = SpreadHeader
                
                    sTmp = "����":      .Col = 8:           .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "�ѱ�":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "�����":    .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��ġ":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "�繮":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "����":      .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                                
            End With
            
            '## ���� �����ο�
            For nCol = 8 To (11 + 8 - 1) Step 1         '<< 11 ����
                nTotinwon = 0
            
                sprSTD.Col = nCol
                sprSTD.Row = SpreadHeader
                    sHeader = Trim(sprSTD.Text)         '<< ����� ������ �����̸� count + 1
                    
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
                
                '## total �ο� üũ
                sprClassDet.Row = 1
                sprClassDet.Col = nCol
                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -9999999, 9999999, ",", nTotinwon)
                
            Next nCol
    
        Case "02"               '<< �ڿ��� : 8 ����
            With sprClassDet
                .MaxCols = 0
                .MaxRows = 0
                
                .MaxCols = 15
                .MaxRows = 2
                
                .Row = SpreadHeader
                
                    sTmp = "��1":       .Col = 8:           .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "ȭ1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��1":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "ȭ2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
                    sTmp = "��2":       .Col = .Col + 1:    .Text = sTmp:       .ColWidth(.Col) = 4
            End With
            
            '## ���� �����ο�
            For nCol = 8 To (8 + 8 - 1) Step 1          '<< 8����
                nTotinwon = 0
            
                sprSTD.Col = nCol
                sprSTD.Row = SpreadHeader
                    sHeader = Trim(sprSTD.Text)         '<< ����� ������ �����̸� count + 1
                    
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
                
                '## total �ο� üũ
                sprClassDet.Row = 1
                sprClassDet.Col = nCol
                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -9999999, 9999999, ",", nTotinwon)
                
            Next nCol
            
    End Select
    
    
    
    
    
'<< ��Ī >>
    With sprClass
        sClass = ""
        nLimit = 0
        nSubj = 0
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 2
                sClass = Trim(.Text)        '<< �� ��Ī
            
            .Col = 5
                nLimit = .value
            
        '>> ���� ���� ���� ����ϰ� ���� �л��� ��û�� ������ �л� ���� ���Ѵ�.
            If Select_Student(sClass, nLimit, nSubj) = True Then

            '>> ������ ���õ� ������ �л��鿡 ���� ��û���� ���� ����.
                Call Select_Order_Gwamok(nSubj)
                
            '>> ������ ���õ� ������ �л����� �ι�°�� ��û�� ���� ������ �л����� ���Ѵ�.
                nTmp = Select_Sec_Order_Gwamok(nLimit, nSubj)
                
            
                If nTmp = 0 Then
                    '>> �ι�° ��û�� ���� �л����� �������� ���� ���
                    Call Make_Class_Less_OrdBok(sClass, nLimit)
                    
                Else
                
                    Call Make_Class_Great_OrdBok(sClass, nLimit, nTmp)
                    
                End If
                
                Call ReAction_sprSTD        '<< �������� ���� �л��� ���� �� �ʱ�ȭ
                
            End If
        Next nRow
    End With
        
        
    With sprClass
        ' ���� �л����� ������.
        For nRow = 1 To .MaxRows Step 1
            nMaxStdinwon = 0
            
            .Row = nRow
            .Col = 2
                sClass = .Text                      '<< �������� ����
            
            For nCol = 6 To .MaxCols Step 1         '<< column ���� ���� : sprClass�� �ѿ�/ ����/ �����ο��� ���� ������ �����Ƿ� column ������ 6����
                
                sTmp = Set_Minus_Class_inwon(sClass, nCol + 2)      '<< �ش� ���� ���� �л��� ���� ����. �����л� ��. (sprClassDet�� ������ 8����, sprClass�� ������ 6����)
                    
                If IsNumeric(sTmp) = True Then
                    nTmp = CLng(sTmp)
                    
                    If nMaxStdinwon <= nTmp Then nMaxStdinwon = nTmp
                    
                    .Col = nCol
                        Call basFunction.Set_SprType_Numeric(sprClass, 0, -999999, 999999, ",", nTmp)
                    
                End If
                
            Next nCol
            
            '## �ִ��ο��� ���� : �����ο�/ �����ο� ���
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


'<< �ش� ���� ���� �л��� ���� ����. �����л� ��.
Private Function Set_Minus_Class_inwon(ByVal aClass As String, ByVal aCol As Long) As Long
    Dim nRow        As Long
    Dim nCnt        As Long
    
    nCnt = 0
    
    With sprSTD
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2
            
            If StrComp(Trim(.Text), aClass, vbTextCompare) = 0 Then             ' �� ���� �л��� �´ٸ�
                .Col = aCol                 ' ���ð����� Ȯ��
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< �ι��� : 11 ����
                        
                        If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then
                           
                            nCnt = nCnt + 1     '<< �����ߴٸ� �����л��� �Ѹ� ����.
                            
                        End If
                        
                    Case "02"
                            
                        If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then
                           
                            nCnt = nCnt + 1     '<< �����ߴٸ� �����л��� �Ѹ� ����.
                            
                        End If
                End Select
                
            End If
        Next nRow
    End With
    
    Set_Minus_Class_inwon = nCnt

End Function


'<< �������� ���� �л��� ���� �� �ʱ�ȭ
Private Sub ReAction_sprSTD()
    Dim nRow        As Long
    Dim nCol        As Long
    
    ' ���� �� �����ϰ� �� �� ������ �л��鿡 ���� ó��.
    ' ������ �л�(maxcols�� 0���� Setting�� �л���)�� �ٽ� �η� �ʱ�ȭ.
    ' maxcols�� 0�̶� ���� ������ ������ �ݿ� �������� ���� �л����� �ǹ���.
    With sprSTD
        .Col = .MaxCols
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols - 2
                If .Text = "0" Then .Text = ""
            
        Next nRow
    End With

End Sub

'>> �ι�° ������ ������ �л����� ���������� ������
'   �ι�° ������ ������ �����Ͽ� �� �л����� ���������� ���������� ����.
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
            
            If StrComp(Trim(.Text), "0", vbTextCompare) = 0 Then            '<< ù��° ���ð����� ���� ������ �����̶��... ���õ� �л����̶��..
                
                .Col = nSubj                '�ι�° ���ð����� ����
                
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< �ι��� : 11 ����
                        
                        If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then
                        
                            nC = nC + 1
                            
                            If nC > nLimit Then Exit Sub
                            
                            .Col = .MaxCols - 2
                                .Text = aClass          '<< "0" �� �ݸ����� ��ġ
                
                            .Col = .MaxCols
                                .value = 0          ' ��������
                            
                            .Row2 = .Row
                            .Col = 1:   .Col2 = .MaxCols
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            
                            For nCol = 8 To (11 + 8 - 1) Step 1
                                .Col = nCol
                                
                                If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then
                                         
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
                            
                        If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                           StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then
                           
                            nC = nC + 1
                            
                            If nC > nLimit Then Exit Sub
                            
                            .Col = .MaxCols - 2
                                .Text = aClass          '<< "0" �� �ݸ����� ��ġ
                
                            .Col = .MaxCols
                                .value = 0          ' ��������
                            
                            .Row2 = .Row
                            .Col = 1:   .Col2 = .MaxCols
                            .BlockMode = True
                                .BackColor = basModule.WhiteColor
                                .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            
                            For nCol = 8 To (8 + 8 - 1) Step 1
                                .Col = nCol
                                
                                If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                                   StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then
                                         
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


'>> �ι�° ���ð������ �ش� ���� �������� ���� ���, ù��° ���ð��� �����л�����
'   ���������� ��������� �߶� �ݿ� �����Ѵ�.
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
                
                .Text = aClass          '<< "0" �� �ݸ����� ��ġ
                
                .Col = .MaxCols
                    .value = 0          ' ��������
                
                .Row2 = .Row
                .Col = 1:   .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                Select Case Trim(Right(txtKaeyol.Text, 30))
                    Case "01", "03"         '<< �ι��� : 11 ����
                        
                        For nCol = 8 To (11 + 8 - 1) Step 1         '< 2007.12.17
                        
                            .Col = nCol
                            
                            If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then
                            
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
                    Case "02"               '<< �ڿ��� : 8 ����
                        
                        For nCol = 8 To (8 + 8 - 1) Step 1          '< 2007.12.17
                        
                            .Col = nCol
                            
                            If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                               StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then
                                                   
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










'>> �л��� �ι�° ������ ���� ������ �����Ѵ�.
'   nLimit �� ���� ���õ� ���� ���� ��, 1�� �� ���ð��� ���� �÷� ��
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
                
            If .Col <> nSubj Then               ' �л��� ù��° ���ð��� ����
                If nLimit < nTmp Then           ' ���� ���� �� ���� ���ٸ� OK(�װ� ����)
                    iSubj = nCol
                    Select_Sec_Order_Gwamok = iSubj
                    
                    Exit Function
                    
                End If
            End If
        Next nCol
    End With
    
    Select_Sec_Order_Gwamok = iSubj
End Function


'>> ������ ���õ� ������ �л��鿡 ���� ��û���� ���� ����.
Private Sub Select_Order_Gwamok(ByVal nSubj As Long)
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sHeader     As String
    Dim sTmp        As String
    Dim nTmp        As Long
    
    With sprClassDet                '<< �������ߴ� ������ ��� �ʱ�ȭ
        .Row = .MaxRows             ' 2��° ���� �ʱ�ȭ
        
        For nCol = 1 To .MaxCols
            .Col = nCol:        .Text = ""
        Next nCol
    End With
    
    ' �� �� �������� ó�� �л����� ����.
    ' maxcols�� �ݸ��� ǥ���ϰ� �Ǿ� �ִµ� �װ� ���̸� ���� ���� �������� �ʾҴٰ� ����.
    ' �ش� ������ ������ �л����� Ȯ���Ѵ�. nsubj�� �÷������� �����ߴµ� ���߿� �ٲ㵵 ��.
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
                
                If StrComp(sHeader, sTmp, vbTextCompare) = 0 Then       ' ������ �����̶��
                
                    .Col = .MaxCols - 2
                    .Text = 0                                           ' �ϴ� ���� �ʱⰪ 0����, 0�� ���õ� �л��̶�� ǥ��
                    
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
                    
                    
                    ' �� ���� ������� ���� �����Ͽ� �ش� �л��� ���� ���ð����� ����Ѵ�.
                    ' �� �� �л��� ������ ����鿡 ���ؼ��� �����Ͽ� sprClassDet�� ������ �ٲ� ��Ȯ�� ������ ����°���.
                    Select Case Trim(Right(txtKaeyol.Text, 30))
                        Case "01", "03"         '<< �ι��� : 11 ����
                            For nCol = 8 To (11 + 8 - 1) Step 1
                                .Col = nCol
                                
                                '���� �����ִ� ����(�Ʊ� �������� �ؼ� ã�� �� ����)�� �ƴ� �ٸ� �����ϰ��,
                                '�׸��� �� �ٸ� ������ �� �л��� ������ �����.. sprClassDet�� ���� ������Ʈ�ؾ���. �����
                            
                                If .Col <> nSubj And _
                                    (StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "����", vbTextCompare) = 0) Then
                                                   
                                        With sprClassDet
                                            .Row = .MaxRows
                                            .Col = nCol
                                            
                                                If IsNumeric(.Text) = False Then
                                                    nTmp = 1
                                                Else
                                                    nTmp = .value + 1
                                                End If
                                                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)    '���ð��� �л� �� ����
                                        End With

                                End If
                                
                            Next nCol
                        
                        Case "02"               '<< �ڿ��� : 8 ����
                        
                            For nCol = 8 To (8 + 8 - 1) Step 1
                                .Col = nCol
                                
                                '���� �����ִ� ����(�Ʊ� �������� �ؼ� ã�� �� ����)�� �ƴ� �ٸ� �����ϰ��,
                                '�׸��� �� �ٸ� ������ �� �л��� ������ �����.. sprClassDet�� ���� ������Ʈ�ؾ���. �����
                            
                                If .Col <> nSubj And _
                                    (StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Or _
                                     StrComp(Trim(.Text), "��2", vbTextCompare) = 0) Then
                                                   
                                        With sprClassDet
                                            .Row = .MaxRows
                                            .Col = nCol
                                            
                                                If IsNumeric(.Text) = False Then
                                                    nTmp = 1
                                                Else
                                                    nTmp = .value + 1
                                                End If
                                                    Call basFunction.Set_SprType_Numeric(sprClassDet, 0, -999999, 999999, ",", nTmp)    '���ð��� �л� �� ����
                                        End With
                                    
                                End If
                                
                            Next nCol
                                
                    End Select
        
                End If
                
            End If
        Next nRow
        
    End With
End Sub


'>> ���� ���� ���� ����ϰ� ���� �л��� ��û�� ������ �л� ���� ���Ѵ�.
Private Function Select_Student(ByVal sBan As String, ByVal nLimit As Integer, ByRef nSubj As Long) As Boolean
    Dim nCols       As Long
    Dim nTmp        As Long
    Dim nC          As Long
    
    Dim bChk        As Boolean

    bChk = False

    nC = 0
    ' sprClass �� �� �л����� �����ִ� ��������.
    ' �� �������忡�� ���� ������ �� nLimit��� ����= ���� ����.
    ' ���� �������� ũ��, �� ū ����� ���� ���� ������ �����Ͽ� �� �ش� �����û�ڸ� ã��
    With sprClassDet
        .Row = 1

        For nCols = 8 To .MaxCols Step 1        ' �� ���� ��ü ������ ��� ����. : ������ 8��° �ٺ��� ����
            .Col = nCols
            
            If IsNumeric(.Text) = True Then
                nTmp = .value
            Else
                nTmp = 0
            End If
            
            If nTmp > nLimit Then               ' ���࿡ ���������� �ش������ ������ �л����� �� ���ٸ� ������
                bChk = True
                
                If nC = 0 Then
                    nC = val(.Text)             ' nC =  �ּ��� ���� ���ϱ� ���� �������.
                                                ' �� ������ ����Ͽ� ������ ��� ���ذ��鼭 �������ٴ� ����, ���� ���� ���� ���� �����ϰ� ��.
                    nSubj = nCols
                    
                ElseIf nC > val(.Text) Then
                    nC = val(.Text)
                    nSubj = nCols
                    
                End If
            End If
        Next nCols
    End With

    If bChk = False Then                ' ���� ���� ���� ������ ���� �ο��� ���ð���ۿ� ������� ó��
                                        ' ���������� �� �۴ٸ�.. ��� �ϸ� ������..
        ' ����� ���� ó������ ����.
    End If

    Select_Student = bChk
End Function
























'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%% �� ����ϱ�
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'>> �ݵ�Ͻ� UPDATE �� �ֽ��ϴ�.


Private Sub cmdinput_Class_Click()

    Dim nRow        As Long
    Dim nChk        As Long
    Dim uClass()    As tClass

    Dim nRec        As Long
    Dim sClassNM    As String
    Dim sTmp        As String

    Dim ninClass()  As Long         ' ������ ��
    Dim nC          As Long

    nChk = 0

    With sprClass
        If .MaxRows = 0 Then
            MsgBox "�� ������ ��ȸ�ϼ���.", vbExclamation + vbOKOnly, "�� ����ϱ�"
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
            .Col = .MaxCols - 2                 '< �ݸ�
            If Trim(.Text) > " " Then
                nChk = nChk + 1
                Exit For
            End If
        Next nRow

        If nChk = 0 Then
            MsgBox "ó���� ���� �����ϴ�.", vbExclamation + vbOKOnly, "�� ����ϱ�"
            Exit Sub
        End If

        ReDim ninClass(0) As Long
        nC = 0

        For nRow = 1 To .MaxRows Step 1
            nChk = 0                        '<< ��ϰ��� üũ

            .Row = nRow
            .Col = .MaxCols - 2             '< �ݸ�
                sClassNM = Trim(.Text)

            If sClassNM > " " Then
                For nRec = 1 To UBound(uClass) Step 1
                    If StrComp(sClassNM, uClass(nRec).CLSNM, vbTextCompare) = 0 Then

                        sTmp = uClass(nRec).CLSCD
                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", LenB(sTmp), sTmp)

                        nChk = nChk + 1

                        '## ������ ��
                        nC = nC + 1
                        ReDim Preserve ninClass(nC) As Long
                        ninClass(nC) = .Row

                    End If
                Next nRec

                If nChk = 0 Then
                    MsgBox Trim(CStr(.Row)) & "��" & vbCrLf & "�� ���� �߸��Ǿ����� Ȯ���Ͻʽÿ�.", vbExclamation + vbOKOnly, "�� ����ϱ�"
                    Exit Sub
                End If
            End If

        Next nRow
    End With

    If UBound(ninClass) > 0 Then
        If input_Class_Data(ninClass) = True Then
            MsgBox "�� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�� ����ϱ�"
        Else
            MsgBox "�� ����� ���Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ����ϱ�"
        End If
    Else
        MsgBox "ó���� ������ �����ϴ�.", vbExclamation + vbOKOnly, "�� ����ϱ�"
    End If

End Sub


'## ����ϱ�
Private Function input_Class_Data(ByRef ainClass() As Long) As Boolean
    Dim bRet        As Boolean

    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
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
            sprSTD.Col = sprSTD.MaxCols - 2             '< �ݸ�
                sClassCD = Trim(sprSTD.Text)

            sprSTD.Row = ainClass(nRec)
            sprSTD.Col = 1                              '< �л��ڵ� (�ý���)
                sSchNO = Trim(sprSTD.Text)

        sStr = ""
        sStr = sStr & " UPDATE CLTTL01TB"
        sStr = sStr & "    SET SEL_CLASS = '" & sClassCD & "'"
        sStr = sStr & "  WHERE SCHNO = '" & sSchNO & "'"
        sStr = sStr & "    AND ACID  = '" & Trim(basModule.SchCD) & "'"

  

'    '>> ���ڵ�
'        sprSTD.Row = ainClass(nRec)
'        sprSTD.Col = sprSTD.MaxCols - 2
'            sClassCD = Trim(sprSTD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SEL_CLASS", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'
'    '>> �л�
'        sprSTD.Row = ainClass(nRec)
'        sprSTD.Col = 1
'            sTmp = Trim(sprSTD.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'
'    '>> �п�
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

        '<< �Ʒ��� �κ��� PROCEDURE ���

 

            '>> �п��ڵ�
            sTmp = Trim(basModule.SchCD)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> ������ ���
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
'** �����ϱ�
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


'<< ���ù� ���� �����ϱ�
Private Sub cmdDeleteClass_Click()
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
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
                sprClass.Col = 1                '< �л��ڵ�
                    sClassCD = Trim(sprClass.Text)

            sStr = ""
            sStr = sStr & " UPDATE CLTTL01TB"
            sStr = sStr & "    SET SEL_CLASS = ''   "       '<< class ����
            sStr = sStr & "  WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "    AND SEL_CLASS = '" & sClassCD & "'"



    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> ��
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
                MsgBox "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
                basDataBase.DBConn.RollbackTrans
                
                Set DBCmd = Nothing
                Set DBParam = Nothing
                
                Call cmdClass_Click         '<< �� ��ȸ
                Call cmdFindStd_Click       '<< �л���ȸ
    
            ElseIf nExe > 0 Then

            '<< �Ʒ��� �κ��� PROCEDURE ���




                '>> �п��ڵ�
                sTmp = Trim(basModule.SchCD)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

                '>> ������ ���
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
        MsgBox "�����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Call cmdClass_Click         '<< �� ��ȸ
    Call cmdFindStd_Click       '<< �л���ȸ

    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub

'<< �����л� �� �����ϱ�
Private Sub cmdDelStdClass_Click()
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
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
                sprSTD.Col = 1              '< �л��ڵ�
                    sSchNO = Trim(sprSTD.Text)

            sStr = ""
            sStr = sStr & " UPDATE CLTTL01TB"
            sStr = sStr & "    SET SEL_CLASS = ''   "       '<< class ����
            sStr = sStr & "  WHERE SCHNO = '" & sSchNO & "'"
            sStr = sStr & "    AND ACID  = '" & Trim(basModule.SchCD) & "'"



    '    '>> �л�
    '        sprSTD.Row = nRec
    '        sprSTD.Col = 1
    '            sTmp = Trim(sprSTD.Text)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("SCHNO", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> �п�
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
                MsgBox "�����л��� �� ������ �����ϴ�.", vbExclamation + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
                basDataBase.DBConn.RollbackTrans
            
                Set DBCmd = Nothing
                Set DBParam = Nothing
                
                Call cmdClass_Click         '<< �� ��ȸ
                Call cmdFindStd_Click       '<< �л���ȸ
                
            ElseIf nExe > 0 Then

            '<< �Ʒ��� �κ��� PROCEDURE ���




                '>> �п��ڵ�
                sTmp = Trim(basModule.SchCD)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

                '>> ������ ���
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
        MsgBox "�����Ͽ����ϴ�.", vbInformation + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    End If

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Call cmdClass_Click         '<< �� ��ȸ
    Call cmdFindStd_Click       '<< �л���ȸ
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "���ù� ��ϳ��� �����ϱ�"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub








'>> ��ü ���� : 2007.12.17
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
                        Case 1                      '<< ���
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 2                      '<< ����
                            .SortKey(nC) = 5
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 3                      '<< �ܱ���
                            .SortKey(nC) = 6
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 4                      '<< �հ�
                            .SortKey(nC) = 7
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 5                      '<< MU_TYPE
                            .SortKey(nC) = .MaxCols - 1
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                        Case 6                      '<< �����ȣ
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
        
        If aClick <> "CMD" Then         '< ��ư Ŭ���� �ƴѰ��
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











