VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form EXM110 
   Caption         =   "�л�����"
   ClientHeight    =   10950
   ClientLeft      =   480
   ClientTop       =   2175
   ClientWidth     =   15135
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15135
   Begin VB.VScrollBar VScroll1 
      Height          =   9855
      Left            =   14250
      TabIndex        =   20
      Top             =   570
      Width           =   225
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14385
         Begin VB.ComboBox cboBan 
            Height          =   300
            Left            =   7110
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   21
            Top             =   60
            Width           =   825
         End
         Begin VB.TextBox txtStdCD 
            Height          =   345
            Left            =   8460
            TabIndex        =   15
            Text            =   "txtStdCD"
            Top             =   60
            Width           =   645
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   9120
            TabIndex        =   14
            Text            =   "txtStdNM"
            Top             =   60
            Width           =   1005
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�л���ȸ(&F)"
            Height          =   375
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   1215
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   5760
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   60
            Width           =   915
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "����page���"
            Height          =   375
            Left            =   10230
            TabIndex        =   6
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "��üpage���"
            Height          =   375
            Left            =   11640
            TabIndex        =   5
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "��"
            Height          =   375
            Left            =   14040
            TabIndex        =   4
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "��"
            Height          =   375
            Left            =   13020
            TabIndex        =   3
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13410
            TabIndex        =   2
            Text            =   "txtPage"
            Top             =   30
            Width           =   615
         End
         Begin EditLib.fpMask fpSTD_Ns 
            Height          =   285
            Left            =   3360
            TabIndex        =   9
            Top             =   75
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   503
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
            Mask            =   "AAAA"
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
         Begin EditLib.fpMask fpSTD_Ne 
            Height          =   285
            Left            =   4470
            TabIndex        =   10
            Top             =   75
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
            _ExtentY        =   503
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
            Mask            =   "AAAA"
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
         Begin EditLib.fpDateTime fpExmYM 
            Height          =   330
            Left            =   1890
            TabIndex        =   17
            Top             =   60
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
            _ExtentY        =   582
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
            AlignTextH      =   0
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
            Text            =   "2004-01"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "YYYY-MM"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   0
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��"
            Height          =   210
            Index           =   4
            Left            =   6090
            TabIndex        =   22
            Top             =   120
            Width           =   975
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����"
            Height          =   210
            Index           =   0
            Left            =   870
            TabIndex        =   18
            Top             =   120
            Width           =   975
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л�"
            Height          =   210
            Index           =   5
            Left            =   7470
            TabIndex        =   16
            Top             =   120
            Width           =   975
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�й�        ����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   13
            Top             =   120
            Width           =   2355
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "��¹�"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   12
            Top             =   30
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�迭"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   5340
            TabIndex        =   11
            Top             =   120
            Width           =   945
         End
      End
   End
   Begin MSComctlLib.ProgressBar progDisp 
      Height          =   135
      Left            =   30
      TabIndex        =   19
      Top             =   450
      Width           =   14445
      _ExtentX        =   25479
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   14130
      Top             =   12780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pReportViewer 
      BackColor       =   &H00FFFFFF&
      Height          =   9795
      Left            =   0
      ScaleHeight     =   9735
      ScaleWidth      =   14295
      TabIndex        =   23
      Top             =   600
      Width           =   14355
      Begin VB.TextBox txtTeacher 
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7230
         TabIndex        =   362
         Text            =   "txtTeacher"
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox txtGaeyol 
         BorderStyle     =   0  '����
         Height          =   285
         Left            =   810
         TabIndex        =   347
         Text            =   "txtGaeyol"
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox txtBan 
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   346
         Text            =   "txtBan"
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox txtStdNM1 
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5070
         TabIndex        =   345
         Text            =   "txtStdNM1"
         Top             =   780
         Width           =   1005
      End
      Begin VB.TextBox txtStdCD1 
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         TabIndex        =   344
         Text            =   "txtStdCD1"
         Top             =   780
         Width           =   615
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   11280
         TabIndex        =   343
         Text            =   "M5"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   12030
         TabIndex        =   342
         Text            =   "M5"
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   11280
         TabIndex        =   341
         Text            =   "M5"
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   11280
         TabIndex        =   340
         Text            =   "M5"
         Top             =   2610
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   11280
         TabIndex        =   339
         Text            =   "M5"
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   11280
         TabIndex        =   338
         Text            =   "M5"
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   11280
         TabIndex        =   337
         Text            =   "M5"
         Top             =   3330
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   11280
         TabIndex        =   336
         Text            =   "M5"
         Top             =   3570
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   11280
         TabIndex        =   335
         Text            =   "M5"
         Top             =   3810
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   11280
         TabIndex        =   334
         Text            =   "M5"
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   11280
         TabIndex        =   333
         Text            =   "M5"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   11280
         TabIndex        =   332
         Text            =   "M5"
         Top             =   4530
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   11280
         TabIndex        =   331
         Text            =   "M5"
         Top             =   4770
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   11280
         TabIndex        =   330
         Text            =   "M5"
         Top             =   5010
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   11280
         TabIndex        =   329
         Text            =   "M5"
         Top             =   5250
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   11280
         TabIndex        =   328
         Text            =   "M5"
         Top             =   5490
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   11280
         TabIndex        =   327
         Text            =   "M5"
         Top             =   5730
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   11280
         TabIndex        =   326
         Text            =   "M5"
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   11280
         TabIndex        =   325
         Text            =   "M5"
         Top             =   6210
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   11280
         TabIndex        =   324
         Text            =   "M5"
         Top             =   6450
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   11280
         TabIndex        =   323
         Text            =   "M5"
         Top             =   6690
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   11280
         TabIndex        =   322
         Text            =   "M5"
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   11280
         TabIndex        =   321
         Text            =   "M5"
         Top             =   7170
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   11280
         TabIndex        =   320
         Text            =   "M5"
         Top             =   7410
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   11280
         TabIndex        =   319
         Text            =   "M5"
         Top             =   7650
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   11280
         TabIndex        =   318
         Text            =   "M5"
         Top             =   7890
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   11280
         TabIndex        =   317
         Text            =   "M5"
         Top             =   8130
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   11280
         TabIndex        =   316
         Text            =   "M5"
         Top             =   8370
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   11280
         TabIndex        =   315
         Text            =   "M5"
         Top             =   8610
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   11280
         TabIndex        =   314
         Text            =   "M5"
         Top             =   8850
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   11280
         TabIndex        =   313
         Text            =   "M5"
         Top             =   9090
         Width           =   1155
      End
      Begin VB.TextBox M5D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   11280
         TabIndex        =   312
         Text            =   "M5"
         Top             =   9330
         Width           =   1155
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   12750
         TabIndex        =   311
         Text            =   "M5"
         Top             =   2130
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   12750
         TabIndex        =   310
         Text            =   "M5"
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   12750
         TabIndex        =   309
         Text            =   "M5"
         Top             =   2610
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   12750
         TabIndex        =   308
         Text            =   "M5"
         Top             =   2850
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   12750
         TabIndex        =   307
         Text            =   "M5"
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   12750
         TabIndex        =   306
         Text            =   "M5"
         Top             =   3330
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   12750
         TabIndex        =   305
         Text            =   "M5"
         Top             =   3570
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   12750
         TabIndex        =   304
         Text            =   "M5"
         Top             =   3810
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   12750
         TabIndex        =   303
         Text            =   "M5"
         Top             =   4050
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   12750
         TabIndex        =   302
         Text            =   "M5"
         Top             =   4290
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   12750
         TabIndex        =   301
         Text            =   "M5"
         Top             =   4530
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   12750
         TabIndex        =   300
         Text            =   "M5"
         Top             =   4770
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   12750
         TabIndex        =   299
         Text            =   "M5"
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   12750
         TabIndex        =   298
         Text            =   "M5"
         Top             =   5250
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   12750
         TabIndex        =   297
         Text            =   "M5"
         Top             =   5490
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   12750
         TabIndex        =   296
         Text            =   "M5"
         Top             =   5730
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   12750
         TabIndex        =   295
         Text            =   "M5"
         Top             =   5970
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   12750
         TabIndex        =   294
         Text            =   "M5"
         Top             =   6210
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   12750
         TabIndex        =   293
         Text            =   "M5"
         Top             =   6450
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   12750
         TabIndex        =   292
         Text            =   "M5"
         Top             =   6690
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   12750
         TabIndex        =   291
         Text            =   "M5"
         Top             =   6930
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   12750
         TabIndex        =   290
         Text            =   "M5"
         Top             =   7170
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   12750
         TabIndex        =   289
         Text            =   "M5"
         Top             =   7410
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   12750
         TabIndex        =   288
         Text            =   "M5"
         Top             =   7650
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   12750
         TabIndex        =   287
         Text            =   "M5"
         Top             =   7890
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   12750
         TabIndex        =   286
         Text            =   "M5"
         Top             =   8130
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   12750
         TabIndex        =   285
         Text            =   "M5"
         Top             =   8370
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   12750
         TabIndex        =   284
         Text            =   "M5"
         Top             =   8610
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   12750
         TabIndex        =   283
         Text            =   "M5"
         Top             =   8850
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   12750
         TabIndex        =   282
         Text            =   "M5"
         Top             =   9090
         Width           =   765
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   12750
         TabIndex        =   281
         Text            =   "M5"
         Top             =   9330
         Width           =   765
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   8700
         TabIndex        =   280
         Text            =   "M4"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9390
         TabIndex        =   279
         Text            =   "M4"
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   8700
         TabIndex        =   278
         Text            =   "M4"
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   8700
         TabIndex        =   277
         Text            =   "M4"
         Top             =   2610
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   8700
         TabIndex        =   276
         Text            =   "M4"
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   8700
         TabIndex        =   275
         Text            =   "M4"
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   8700
         TabIndex        =   274
         Text            =   "M4"
         Top             =   3330
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   8700
         TabIndex        =   273
         Text            =   "M4"
         Top             =   3570
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   8700
         TabIndex        =   272
         Text            =   "M4"
         Top             =   3810
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   8700
         TabIndex        =   271
         Text            =   "M4"
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   8700
         TabIndex        =   270
         Text            =   "M4"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   8700
         TabIndex        =   269
         Text            =   "M4"
         Top             =   4530
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   8700
         TabIndex        =   268
         Text            =   "M4"
         Top             =   4770
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   8700
         TabIndex        =   267
         Text            =   "M4"
         Top             =   5010
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   8700
         TabIndex        =   266
         Text            =   "M4"
         Top             =   5250
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   8700
         TabIndex        =   265
         Text            =   "M4"
         Top             =   5490
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   8700
         TabIndex        =   264
         Text            =   "M4"
         Top             =   5730
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   8700
         TabIndex        =   263
         Text            =   "M4"
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   8700
         TabIndex        =   262
         Text            =   "M4"
         Top             =   6210
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   8700
         TabIndex        =   261
         Text            =   "M4"
         Top             =   6450
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   8700
         TabIndex        =   260
         Text            =   "M4"
         Top             =   6690
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   8700
         TabIndex        =   259
         Text            =   "M4"
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   8700
         TabIndex        =   258
         Text            =   "M4"
         Top             =   7170
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   8700
         TabIndex        =   257
         Text            =   "M4"
         Top             =   7410
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   8700
         TabIndex        =   256
         Text            =   "M4"
         Top             =   7650
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   8700
         TabIndex        =   255
         Text            =   "M4"
         Top             =   7890
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   8700
         TabIndex        =   254
         Text            =   "M4"
         Top             =   8130
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   8700
         TabIndex        =   253
         Text            =   "M4"
         Top             =   8370
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   8700
         TabIndex        =   252
         Text            =   "M4"
         Top             =   8610
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   8700
         TabIndex        =   251
         Text            =   "M4"
         Top             =   8850
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   8700
         TabIndex        =   250
         Text            =   "M4"
         Top             =   9090
         Width           =   1155
      End
      Begin VB.TextBox M4D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   8700
         TabIndex        =   249
         Text            =   "M4"
         Top             =   9330
         Width           =   1155
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   10200
         TabIndex        =   248
         Text            =   "M4"
         Top             =   2130
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   10200
         TabIndex        =   247
         Text            =   "M4"
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   10200
         TabIndex        =   246
         Text            =   "M4"
         Top             =   2610
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   10200
         TabIndex        =   245
         Text            =   "M4"
         Top             =   2850
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   10200
         TabIndex        =   244
         Text            =   "M4"
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   10200
         TabIndex        =   243
         Text            =   "M4"
         Top             =   3330
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   10200
         TabIndex        =   242
         Text            =   "M4"
         Top             =   3570
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   10200
         TabIndex        =   241
         Text            =   "M4"
         Top             =   3810
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   10200
         TabIndex        =   240
         Text            =   "M4"
         Top             =   4050
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   10200
         TabIndex        =   239
         Text            =   "M4"
         Top             =   4290
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   10200
         TabIndex        =   238
         Text            =   "M4"
         Top             =   4530
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   10200
         TabIndex        =   237
         Text            =   "M4"
         Top             =   4770
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   10200
         TabIndex        =   236
         Text            =   "M4"
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   10200
         TabIndex        =   235
         Text            =   "M4"
         Top             =   5250
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   10200
         TabIndex        =   234
         Text            =   "M4"
         Top             =   5490
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   10200
         TabIndex        =   233
         Text            =   "M4"
         Top             =   5730
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   10200
         TabIndex        =   232
         Text            =   "M4"
         Top             =   5970
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   10200
         TabIndex        =   231
         Text            =   "M4"
         Top             =   6210
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   10200
         TabIndex        =   230
         Text            =   "M4"
         Top             =   6450
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   10200
         TabIndex        =   229
         Text            =   "M4"
         Top             =   6690
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   10200
         TabIndex        =   228
         Text            =   "M4"
         Top             =   6930
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   10200
         TabIndex        =   227
         Text            =   "M4"
         Top             =   7170
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   10200
         TabIndex        =   226
         Text            =   "M4"
         Top             =   7410
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   10200
         TabIndex        =   225
         Text            =   "M4"
         Top             =   7650
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   10200
         TabIndex        =   224
         Text            =   "M4"
         Top             =   7890
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   10200
         TabIndex        =   223
         Text            =   "M4"
         Top             =   8130
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   10200
         TabIndex        =   222
         Text            =   "M4"
         Top             =   8370
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   10200
         TabIndex        =   221
         Text            =   "M4"
         Top             =   8610
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   10200
         TabIndex        =   220
         Text            =   "M4"
         Top             =   8850
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   10200
         TabIndex        =   219
         Text            =   "M4"
         Top             =   9090
         Width           =   765
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   10200
         TabIndex        =   218
         Text            =   "M4"
         Top             =   9330
         Width           =   765
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   217
         Text            =   "M3"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6900
         TabIndex        =   216
         Text            =   "M3"
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6240
         TabIndex        =   215
         Text            =   "M3"
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   214
         Text            =   "M3"
         Top             =   2610
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   6240
         TabIndex        =   213
         Text            =   "M3"
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   6240
         TabIndex        =   212
         Text            =   "M3"
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6240
         TabIndex        =   211
         Text            =   "M3"
         Top             =   3330
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   6240
         TabIndex        =   210
         Text            =   "M3"
         Top             =   3570
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   6240
         TabIndex        =   209
         Text            =   "M3"
         Top             =   3810
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   6240
         TabIndex        =   208
         Text            =   "M3"
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   6240
         TabIndex        =   207
         Text            =   "M3"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   6240
         TabIndex        =   206
         Text            =   "M3"
         Top             =   4530
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   6240
         TabIndex        =   205
         Text            =   "M3"
         Top             =   4770
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   6240
         TabIndex        =   204
         Text            =   "M3"
         Top             =   5010
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   6240
         TabIndex        =   203
         Text            =   "M3"
         Top             =   5250
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   6240
         TabIndex        =   202
         Text            =   "M3"
         Top             =   5490
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   6240
         TabIndex        =   201
         Text            =   "M3"
         Top             =   5730
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   6240
         TabIndex        =   200
         Text            =   "M3"
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   6240
         TabIndex        =   199
         Text            =   "M3"
         Top             =   6210
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   6240
         TabIndex        =   198
         Text            =   "M3"
         Top             =   6450
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   6240
         TabIndex        =   197
         Text            =   "M3"
         Top             =   6690
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   6240
         TabIndex        =   196
         Text            =   "M3"
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   6240
         TabIndex        =   195
         Text            =   "M3"
         Top             =   7170
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   6240
         TabIndex        =   194
         Text            =   "M3"
         Top             =   7410
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   6240
         TabIndex        =   193
         Text            =   "M3"
         Top             =   7650
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   6240
         TabIndex        =   192
         Text            =   "M3"
         Top             =   7890
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   6240
         TabIndex        =   191
         Text            =   "M3"
         Top             =   8130
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   6240
         TabIndex        =   190
         Text            =   "M3"
         Top             =   8370
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   6240
         TabIndex        =   189
         Text            =   "M3"
         Top             =   8610
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   6240
         TabIndex        =   188
         Text            =   "M3"
         Top             =   8850
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   6240
         TabIndex        =   187
         Text            =   "M3"
         Top             =   9090
         Width           =   1155
      End
      Begin VB.TextBox M3D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   6240
         TabIndex        =   186
         Text            =   "M3"
         Top             =   9330
         Width           =   1155
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7680
         TabIndex        =   185
         Text            =   "M3"
         Top             =   2130
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7680
         TabIndex        =   184
         Text            =   "M3"
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   7680
         TabIndex        =   183
         Text            =   "M3"
         Top             =   2610
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   7680
         TabIndex        =   182
         Text            =   "M3"
         Top             =   2850
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   7680
         TabIndex        =   181
         Text            =   "M3"
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   7680
         TabIndex        =   180
         Text            =   "M3"
         Top             =   3330
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   7680
         TabIndex        =   179
         Text            =   "M3"
         Top             =   3570
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   7680
         TabIndex        =   178
         Text            =   "M3"
         Top             =   3810
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   7680
         TabIndex        =   177
         Text            =   "M3"
         Top             =   4050
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   7680
         TabIndex        =   176
         Text            =   "M3"
         Top             =   4290
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   7680
         TabIndex        =   175
         Text            =   "M3"
         Top             =   4530
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   7680
         TabIndex        =   174
         Text            =   "M3"
         Top             =   4770
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   7680
         TabIndex        =   173
         Text            =   "M3"
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   7680
         TabIndex        =   172
         Text            =   "M3"
         Top             =   5250
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   7680
         TabIndex        =   171
         Text            =   "M3"
         Top             =   5490
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   7680
         TabIndex        =   170
         Text            =   "M3"
         Top             =   5730
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   7680
         TabIndex        =   169
         Text            =   "M3"
         Top             =   5970
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   7680
         TabIndex        =   168
         Text            =   "M3"
         Top             =   6210
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   7680
         TabIndex        =   167
         Text            =   "M3"
         Top             =   6450
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   7680
         TabIndex        =   166
         Text            =   "M3"
         Top             =   6690
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   7680
         TabIndex        =   165
         Text            =   "M3"
         Top             =   6930
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   7680
         TabIndex        =   164
         Text            =   "M3"
         Top             =   7170
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   7680
         TabIndex        =   163
         Text            =   "M3"
         Top             =   7410
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   7680
         TabIndex        =   162
         Text            =   "M3"
         Top             =   7650
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   7680
         TabIndex        =   161
         Text            =   "M3"
         Top             =   7890
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   7680
         TabIndex        =   160
         Text            =   "M3"
         Top             =   8130
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   7680
         TabIndex        =   159
         Text            =   "M3"
         Top             =   8370
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   7680
         TabIndex        =   158
         Text            =   "M3"
         Top             =   8610
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   7680
         TabIndex        =   157
         Text            =   "M3"
         Top             =   8850
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   7680
         TabIndex        =   156
         Text            =   "M3"
         Top             =   9090
         Width           =   765
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   7680
         TabIndex        =   155
         Text            =   "M3"
         Top             =   9330
         Width           =   765
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   154
         Text            =   "M2"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   4500
         TabIndex        =   153
         Text            =   "M2"
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   152
         Text            =   "M2"
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   151
         Text            =   "M2"
         Top             =   2610
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   150
         Text            =   "M2"
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   149
         Text            =   "M2"
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   148
         Text            =   "M2"
         Top             =   3330
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   3840
         TabIndex        =   147
         Text            =   "M2"
         Top             =   3570
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   146
         Text            =   "M2"
         Top             =   3810
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3840
         TabIndex        =   145
         Text            =   "M2"
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   3840
         TabIndex        =   144
         Text            =   "M2"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3840
         TabIndex        =   143
         Text            =   "M2"
         Top             =   4530
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3840
         TabIndex        =   142
         Text            =   "M2"
         Top             =   4770
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   3840
         TabIndex        =   141
         Text            =   "M2"
         Top             =   5010
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   3840
         TabIndex        =   140
         Text            =   "M2"
         Top             =   5250
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   3840
         TabIndex        =   139
         Text            =   "M2"
         Top             =   5490
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   3840
         TabIndex        =   138
         Text            =   "M2"
         Top             =   5730
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   3840
         TabIndex        =   137
         Text            =   "M2"
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   3840
         TabIndex        =   136
         Text            =   "M2"
         Top             =   6210
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   3840
         TabIndex        =   135
         Text            =   "M2"
         Top             =   6450
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   3840
         TabIndex        =   134
         Text            =   "M2"
         Top             =   6690
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   3840
         TabIndex        =   133
         Text            =   "M2"
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   3840
         TabIndex        =   132
         Text            =   "M2"
         Top             =   7170
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   3840
         TabIndex        =   131
         Text            =   "M2"
         Top             =   7410
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   3840
         TabIndex        =   130
         Text            =   "M2"
         Top             =   7650
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   3840
         TabIndex        =   129
         Text            =   "M2"
         Top             =   7890
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   3840
         TabIndex        =   128
         Text            =   "M2"
         Top             =   8130
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   3840
         TabIndex        =   127
         Text            =   "M2"
         Top             =   8370
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   3840
         TabIndex        =   126
         Text            =   "M2"
         Top             =   8610
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   3840
         TabIndex        =   125
         Text            =   "M2"
         Top             =   8850
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   3840
         TabIndex        =   124
         Text            =   "M2"
         Top             =   9090
         Width           =   1155
      End
      Begin VB.TextBox M2D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   3840
         TabIndex        =   123
         Text            =   "M2"
         Top             =   9330
         Width           =   1155
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5250
         TabIndex        =   122
         Text            =   "M2"
         Top             =   2130
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5250
         TabIndex        =   121
         Text            =   "M2"
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5250
         TabIndex        =   120
         Text            =   "M2"
         Top             =   2610
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5250
         TabIndex        =   119
         Text            =   "M2"
         Top             =   2850
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   5250
         TabIndex        =   118
         Text            =   "M2"
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5250
         TabIndex        =   117
         Text            =   "M2"
         Top             =   3330
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   5250
         TabIndex        =   116
         Text            =   "M2"
         Top             =   3570
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   5250
         TabIndex        =   115
         Text            =   "M2"
         Top             =   3810
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   5250
         TabIndex        =   114
         Text            =   "M2"
         Top             =   4050
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   5250
         TabIndex        =   113
         Text            =   "M2"
         Top             =   4290
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   5250
         TabIndex        =   112
         Text            =   "M2"
         Top             =   4530
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   5250
         TabIndex        =   111
         Text            =   "M2"
         Top             =   4770
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   5250
         TabIndex        =   110
         Text            =   "M2"
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   5250
         TabIndex        =   109
         Text            =   "M2"
         Top             =   5250
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   5250
         TabIndex        =   108
         Text            =   "M2"
         Top             =   5490
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   5250
         TabIndex        =   107
         Text            =   "M2"
         Top             =   5730
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   5250
         TabIndex        =   106
         Text            =   "M2"
         Top             =   5970
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   5250
         TabIndex        =   105
         Text            =   "M2"
         Top             =   6210
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   5250
         TabIndex        =   104
         Text            =   "M2"
         Top             =   6450
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   5250
         TabIndex        =   103
         Text            =   "M2"
         Top             =   6690
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   5250
         TabIndex        =   102
         Text            =   "M2"
         Top             =   6930
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   5250
         TabIndex        =   101
         Text            =   "M2"
         Top             =   7170
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   5250
         TabIndex        =   100
         Text            =   "M2"
         Top             =   7410
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   5250
         TabIndex        =   99
         Text            =   "M2"
         Top             =   7650
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   5250
         TabIndex        =   98
         Text            =   "M2"
         Top             =   7890
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   5250
         TabIndex        =   97
         Text            =   "M2"
         Top             =   8130
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   5250
         TabIndex        =   96
         Text            =   "M2"
         Top             =   8370
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   5250
         TabIndex        =   95
         Text            =   "M2"
         Top             =   8610
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   5250
         TabIndex        =   94
         Text            =   "M2"
         Top             =   8850
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   5250
         TabIndex        =   93
         Text            =   "M2"
         Top             =   9090
         Width           =   765
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   5250
         TabIndex        =   92
         Text            =   "M2"
         Top             =   9330
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   2910
         TabIndex        =   91
         Text            =   "M1"
         Top             =   9330
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   2910
         TabIndex        =   90
         Text            =   "M1"
         Top             =   9090
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   2910
         TabIndex        =   89
         Text            =   "M1"
         Top             =   8850
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   2910
         TabIndex        =   88
         Text            =   "M1"
         Top             =   8610
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   2910
         TabIndex        =   87
         Text            =   "M1"
         Top             =   8370
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   2910
         TabIndex        =   86
         Text            =   "M1"
         Top             =   8130
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   2910
         TabIndex        =   85
         Text            =   "M1"
         Top             =   7890
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   2910
         TabIndex        =   84
         Text            =   "M1"
         Top             =   7650
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   2910
         TabIndex        =   83
         Text            =   "M1"
         Top             =   7410
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   2910
         TabIndex        =   82
         Text            =   "M1"
         Top             =   7170
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   2910
         TabIndex        =   81
         Text            =   "M1"
         Top             =   6930
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   2910
         TabIndex        =   80
         Text            =   "M1"
         Top             =   6690
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   2910
         TabIndex        =   79
         Text            =   "M1"
         Top             =   6450
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   2910
         TabIndex        =   78
         Text            =   "M1"
         Top             =   6210
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   2910
         TabIndex        =   77
         Text            =   "M1"
         Top             =   5970
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   2910
         TabIndex        =   76
         Text            =   "M1"
         Top             =   5730
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   2910
         TabIndex        =   75
         Text            =   "M1"
         Top             =   5490
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   2910
         TabIndex        =   74
         Text            =   "M1"
         Top             =   5250
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   2910
         TabIndex        =   73
         Text            =   "M1"
         Top             =   5010
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   2910
         TabIndex        =   72
         Text            =   "M1"
         Top             =   4770
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   2910
         TabIndex        =   71
         Text            =   "M1"
         Top             =   4530
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   2910
         TabIndex        =   70
         Text            =   "M1"
         Top             =   4290
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   2910
         TabIndex        =   69
         Text            =   "M1"
         Top             =   4050
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   2910
         TabIndex        =   68
         Text            =   "M1"
         Top             =   3810
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   2910
         TabIndex        =   67
         Text            =   "M1"
         Top             =   3570
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2910
         TabIndex        =   66
         Text            =   "M1"
         Top             =   3330
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2910
         TabIndex        =   65
         Text            =   "M1"
         Top             =   3090
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2910
         TabIndex        =   64
         Text            =   "M1"
         Top             =   2850
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2910
         TabIndex        =   63
         Text            =   "M1"
         Top             =   2610
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2910
         TabIndex        =   62
         Text            =   "M1"
         Top             =   2370
         Width           =   765
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2910
         TabIndex        =   61
         Text            =   "M1"
         Top             =   2130
         Width           =   765
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   1440
         TabIndex        =   60
         Text            =   "M1"
         Top             =   9330
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   1440
         TabIndex        =   59
         Text            =   "M1"
         Top             =   9090
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   29
         Left            =   1440
         TabIndex        =   58
         Text            =   "M1"
         Top             =   8850
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   1440
         TabIndex        =   57
         Text            =   "M1"
         Top             =   8610
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   1440
         TabIndex        =   56
         Text            =   "M1"
         Top             =   8370
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   1440
         TabIndex        =   55
         Text            =   "M1"
         Top             =   8130
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   1440
         TabIndex        =   54
         Text            =   "M1"
         Top             =   7890
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   1440
         TabIndex        =   53
         Text            =   "M1"
         Top             =   7650
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   1440
         TabIndex        =   52
         Text            =   "M1"
         Top             =   7410
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   1440
         TabIndex        =   51
         Text            =   "M1"
         Top             =   7170
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   1440
         TabIndex        =   50
         Text            =   "M1"
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   1440
         TabIndex        =   49
         Text            =   "M1"
         Top             =   6690
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   1440
         TabIndex        =   48
         Text            =   "M1"
         Top             =   6450
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   1440
         TabIndex        =   47
         Text            =   "M1"
         Top             =   6210
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   1440
         TabIndex        =   46
         Text            =   "M1"
         Top             =   5970
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   1440
         TabIndex        =   45
         Text            =   "M1"
         Top             =   5730
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   1440
         TabIndex        =   44
         Text            =   "M1"
         Top             =   5490
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   1440
         TabIndex        =   43
         Text            =   "M1"
         Top             =   5250
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   1440
         TabIndex        =   42
         Text            =   "M1"
         Top             =   5010
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   1440
         TabIndex        =   41
         Text            =   "M1"
         Top             =   4770
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1440
         TabIndex        =   40
         Text            =   "M1"
         Top             =   4530
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1440
         TabIndex        =   39
         Text            =   "M1"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1440
         TabIndex        =   38
         Text            =   "M1"
         Top             =   4050
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1440
         TabIndex        =   37
         Text            =   "M1"
         Top             =   3810
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1440
         TabIndex        =   36
         Text            =   "M1"
         Top             =   3570
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1440
         TabIndex        =   35
         Text            =   "M1"
         Top             =   3330
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   34
         Text            =   "M1"
         Top             =   3090
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   33
         Text            =   "M1"
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   32
         Text            =   "M1"
         Top             =   2610
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   31
         Text            =   "M1"
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   30
         Text            =   "M1"
         Top             =   1440
         Width           =   1305
      End
      Begin VB.TextBox M1D 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   29
         Text            =   "M1"
         Top             =   2130
         Width           =   1155
      End
      Begin VB.TextBox M1N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   28
         Text            =   "M1"
         Top             =   1440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox M2N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   27
         Text            =   "M2"
         Top             =   1440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox M3N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7410
         TabIndex        =   26
         Text            =   "M3"
         Top             =   1440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox M4N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9930
         TabIndex        =   25
         Text            =   "M4"
         Top             =   1440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox M5N 
         Appearance      =   0  '���
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   12630
         TabIndex        =   24
         Text            =   "M5"
         Top             =   1440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   24
         X1              =   2880
         X2              =   3630
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   23
         X1              =   4950
         X2              =   6240
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   22
         X1              =   7110
         X2              =   8400
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Image Photo 
         Height          =   435
         Left            =   10980
         Picture         =   "EXM110.frx":0000
         Stretch         =   -1  'True
         Top             =   630
         Width           =   2595
      End
      Begin VB.Label Label21 
         BackStyle       =   0  '����
         Caption         =   "6"
         Height          =   210
         Left            =   960
         TabIndex        =   369
         Top             =   8640
         Width           =   315
      End
      Begin VB.Label Label20 
         BackStyle       =   0  '����
         Caption         =   "5"
         Height          =   210
         Left            =   960
         TabIndex        =   368
         Top             =   7440
         Width           =   315
      End
      Begin VB.Label Label19 
         BackStyle       =   0  '����
         Caption         =   "4"
         Height          =   210
         Left            =   960
         TabIndex        =   367
         Top             =   6240
         Width           =   315
      End
      Begin VB.Label Label18 
         BackStyle       =   0  '����
         Caption         =   "3"
         Height          =   210
         Left            =   960
         TabIndex        =   366
         Top             =   5040
         Width           =   315
      End
      Begin VB.Label Label17 
         BackStyle       =   0  '����
         Caption         =   "2"
         Height          =   210
         Left            =   960
         TabIndex        =   365
         Top             =   3840
         Width           =   315
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '����
         Caption         =   "1"
         Height          =   210
         Left            =   960
         TabIndex        =   364
         Top             =   2670
         Width           =   315
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   21
         X1              =   570
         X2              =   13680
         Y1              =   9300
         Y2              =   9300
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   20
         X1              =   600
         X2              =   13710
         Y1              =   8100
         Y2              =   8100
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   19
         X1              =   600
         X2              =   13710
         Y1              =   6900
         Y2              =   6900
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   18
         X1              =   570
         X2              =   13680
         Y1              =   5700
         Y2              =   5700
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   17
         X1              =   570
         X2              =   13680
         Y1              =   4500
         Y2              =   4500
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  '��
         Index           =   15
         X1              =   570
         X2              =   13680
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   6660
         TabIndex        =   363
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label16 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   12780
         TabIndex        =   361
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '����
         Caption         =   "��¥"
         Height          =   210
         Left            =   11610
         TabIndex        =   360
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   10260
         TabIndex        =   359
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '����
         Caption         =   "��¥"
         Height          =   210
         Left            =   9060
         TabIndex        =   358
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   7710
         TabIndex        =   357
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '����
         Caption         =   "��¥"
         Height          =   210
         Left            =   6570
         TabIndex        =   356
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   5340
         TabIndex        =   355
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '����
         Caption         =   "��¥"
         Height          =   210
         Left            =   4290
         TabIndex        =   354
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   2970
         TabIndex        =   353
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '����
         Caption         =   "��¥"
         Height          =   210
         Left            =   1920
         TabIndex        =   352
         Top             =   1740
         Width           =   495
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   14
         X1              =   570
         X2              =   1380
         Y1              =   1350
         Y2              =   1980
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   13
         X1              =   1380
         X2              =   13680
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   12
         X1              =   12450
         X2              =   12450
         Y1              =   1710
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   11
         X1              =   9870
         X2              =   9870
         Y1              =   1710
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   10
         X1              =   7410
         X2              =   7410
         Y1              =   1710
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   9
         X1              =   5010
         X2              =   5010
         Y1              =   1710
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   8
         X1              =   2640
         X2              =   2640
         Y1              =   1710
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   7
         X1              =   13710
         X2              =   13710
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   6
         X1              =   11160
         X2              =   11160
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   5
         X1              =   8610
         X2              =   8610
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   4
         X1              =   6150
         X2              =   6150
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   3
         X1              =   3780
         X2              =   3780
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   2
         X1              =   570
         X2              =   13680
         Y1              =   9660
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   16
         X1              =   1380
         X2              =   1380
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   1
         X1              =   570
         X2              =   13680
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   0
         X1              =   570
         X2              =   570
         Y1              =   1350
         Y2              =   9660
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Index           =   28
         X1              =   570
         X2              =   13710
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '����
         Caption         =   "����"
         Height          =   210
         Left            =   4530
         TabIndex        =   351
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '����
         Caption         =   "��"
         Height          =   210
         Left            =   3720
         TabIndex        =   350
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '����
         Caption         =   "��"
         Height          =   210
         Left            =   2370
         TabIndex        =   349
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '����
         Caption         =   "�� �� �ܾ���� ����ǥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5070
         TabIndex        =   348
         Top             =   30
         Width           =   4605
      End
   End
End
Attribute VB_Name = "EXM110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type tSTD
    
    ACID        As String
    GAEYOL      As String
    STDCD       As String
    STDNM       As String
    BAN         As String
    
    M1D(31)     As String
    M1N(31)     As String
    
    M2D(31)     As String
    M2N(31)     As String
    
    M3D(31)     As String
    M3N(31)     As String
    
    M4D(31)     As String
    M4N(31)     As String
    
    M5D(31)     As String
    M5N(31)     As String
    
End Type
Private uSTD()      As tSTD
Private nTotRec     As Long         '<< ��ü �л���

Private Sub Form_Load()
        
    Dim nC      As Integer
    
    Me.Tag = "LOAD"
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    fpExmYM.Text = Format(Now, "YYYY-MM-DD")
    
    fpSTD_Ns.Text = ""
    fpSTD_Ne.Text = ""
    
    txtStdCD.Text = ""
    txtStdNM.Text = ""
    
    txtPage.Text = ""
    txtTeacher.Text = ""
    
    
    cboKaeyol.Clear
    '>> �迭
        With cboKaeyol
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
'        '<< �迭 >> : 2008.01.09
'            If Trim(basModule.SchCD) = "N" Then             '< �뷮��
'                .AddItem "��ü" & Space(30) & "03"
'                .AddItem "����(��)" & Space(30) & "04"
'                .AddItem "�ι�����" & Space(30) & "05"
'                .AddItem "�ڿ�����" & Space(30) & "06"
'
'                .AddItem "�ι�-��" & Space(30) & "07"
'                .AddItem "�ڿ�-��" & Space(30) & "08"
'                '.AddItem "�����ι�-��" & Space(30) & "09"
'                '.AddItem "�����ڿ�-��" & Space(30) & "10"
'
'                .AddItem "��)�ι�" & Space(30) & "11"
'                .AddItem "��)�ڿ�" & Space(30) & "12"
'                .AddItem "��)��ü" & Space(30) & "13"
'                .AddItem "��)����(��)" & Space(30) & "14"
'                .AddItem "��)�ι�����" & Space(30) & "15"
'                .AddItem "��)�ڿ�����" & Space(30) & "16"
'            End If
'        '<< �迭 >> : 2008.01.10
'            If Trim(basModule.SchCD) = "K" Then             '< ����
'                .AddItem "�ָ�����" & Space(30) & "04"
'                .AddItem "�ָ��Ǵ�" & Space(30) & "05"
'
'                .AddItem "�߰�����" & Space(30) & "06"
'                .AddItem "�߰��Ǵ�" & Space(30) & "07"
'
'                .AddItem "�������ι�" & Space(30) & "11"
'                .AddItem "�������ڿ�" & Space(30) & "12"
'
'                .AddItem "�������ι�16" & Space(30) & "16"
'                .AddItem "�������ڿ�17" & Space(30) & "17"
'
'            End If
'        '<< �迭 >> : 2009.01.08
'            Select Case Trim(basModule.SchCD)
'                Case "S", "P"
'''                    .AddItem "��ü��" & Space(30) & "03"
'''
'''                    .AddItem "�����ι�" & Space(30) & "05"
'''                    .AddItem "�����ڿ�" & Space(30) & "06"
'
'                    .AddItem "�ι������̾�" & Space(30) & "18"
'                    .AddItem "�ڿ������̾�" & Space(30) & "19"
'
'            End Select
'
'            Select Case Trim(basModule.SchCD)
'                Case "J"
'                    .AddItem "��ü��" & Space(30) & "03"
'
'                    .AddItem "�ż��ι�" & Space(30) & "11"
'                    .AddItem "�ż��ڿ�" & Space(30) & "12"
'
'                    .AddItem "�ι������̾�" & Space(30) & "18"
'                    .AddItem "�ڿ������̾�" & Space(30) & "19"
'
'            End Select
'
'        '<< �迭 >> : 2009.01.09
'            If Trim(basModule.SchCD) = "B" Then             '< �λ�
'
'                .AddItem "���м����ι�" & Space(30) & "05"
'                .AddItem "���м����ڿ�" & Space(30) & "06"
'
'                .AddItem "��.�����ι�" & Space(30) & "07"
'                .AddItem "��.�����ڿ�" & Space(30) & "08"
'
'                .AddItem "��ȭ�ι�" & Space(30) & "09"
'                .AddItem "��ȭ�ڿ�" & Space(30) & "10"
'
'            End If
            
            .ListIndex = 0
        End With
    
    For nC = 1 To 31 Step 1
        M1D(nC).Text = ""
        M1N(nC).Text = ""
        
        M2D(nC).Text = ""
        M2N(nC).Text = ""
        
        M3D(nC).Text = ""
        M3N(nC).Text = ""
        
        M4D(nC).Text = ""
        M4N(nC).Text = ""
        
        M5D(nC).Text = ""
        M5N(nC).Text = ""
    Next nC
    
    Call inits
    
    progDisp.Max = 100
    progDisp.Min = 0
    progDisp.Value = 0
    progDisp.Visible = False
    
    VScroll1.Min = 1
    VScroll1.Max = 100
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 1
    VScroll1.Enabled = False
    
    Me.Tag = ""

End Sub

Private Sub inits()
    Dim nDay            As Integer
    
    '>> �ʱ�ȭ
    txtGaeyol.Text = ""
    txtBan.Text = ""
    txtStdCD1.Text = ""
    txtStdNM1.Text = ""
    
    
    For nDay = 0 To 31 Step 1
        M1D(nDay).FontBold = False:         M1D(nDay).Text = ""
        M1N(nDay).FontBold = False:         M1N(nDay).Text = ""
        
        M2D(nDay).FontBold = False:         M2D(nDay).Text = ""
        M2N(nDay).FontBold = False:         M2N(nDay).Text = ""
        
        M3D(nDay).FontBold = False:         M3D(nDay).Text = ""
        M3N(nDay).FontBold = False:         M3N(nDay).Text = ""
        
        M4D(nDay).FontBold = False:         M4D(nDay).Text = ""
        M4N(nDay).FontBold = False:         M4N(nDay).Text = ""
        
        M5D(nDay).FontBold = False:         M5D(nDay).Text = ""
        M5N(nDay).FontBold = False:         M5N(nDay).Text = ""
    Next nDay
    
End Sub

'## �迭��ȸ
Private Sub cboKaeyol_Click()
    '�ش� �迭�� ����ȸ
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim nLength     As Long
    
    Dim sStr        As String
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim ni          As Integer
    Dim nRec        As Long
    Dim nColor      As Long
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT BAN"
    sStr = sStr & "    FROM SDEXM10TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(Right(cboKaeyol.Text, 10)) <> "ALL" Then
        sStr = sStr & " AND GAEYOL = '" & Trim(Right(cboKaeyol.Text, 10)) & "'"
    End If
    sStr = sStr & "   GROUP BY BAN "
    sStr = sStr & "   ORDER BY BAN "
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    cboBan.Clear
    cboBan.AddItem "��ü" & Space(30) & "ALL"
            
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                
                If IsNull(.Fields("BAN")) = False Then
                    sTmp = Trim(.Fields("BAN"))
                Else
                    sTmp = ""
                End If
                
                cboBan.AddItem sTmp
                
                .MoveNext
            Next nRec
            
        End If
    End With
    
    
    If cboBan.ListCount > 0 Then cboBan.ListIndex = 0
    
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�� ��ȸ"
End Sub


Private Sub txtStdCD_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        If Trim(txtStdCD.Text) > " " Then
            txtStdNM.Text = Find_StdNM(txtStdCD.Text)
            
        End If
    End If
End Sub

Private Sub txtStdNM_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        If Trim(txtStdNM.Text) > " " Then
            txtStdCD.Text = Find_StdCD(txtStdNM.Text, txtStdNM)
            
        End If
    End If
    
End Sub


Private Function Find_StdCD(ByVal aStdNM As String, ByRef aObj As Object) As String
    
    '�ش� �迭�� ����ȸ
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim sStr        As String
    Dim sStdCD      As String
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT STDCD, STDNM "
    sStr = sStr & "    FROM SDEXM10TB"
    sStr = sStr & "   WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND STDNM LIKE '%" & Trim(aStdNM) & "%'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            sStdCD = "":
            If IsNull(.Fields("STDCD")) = False Then
                sStdCD = Trim(.Fields("STDCD"))
                aObj.Text = Trim(.Fields("STDNM"))
            Else
                sStdCD = ""
            End If
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Find_StdCD = sStdCD
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�л���ȣ ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
    
    Find_StdCD = ""
    
End Function


Private Function Find_StdNM(ByVal aStdCD As String) As String
    
    '�ش� �迭�� ����ȸ
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim sStr        As String
    Dim sStdCD      As String
    Dim sRet        As String
    
    On Error GoTo ErrStmt
    
    sStdCD = Format(CLng(aStdCD), "0000")
    
    sStr = ""
    sStr = sStr & "  SELECT STDNM"
    sStr = sStr & "    FROM SDEXM10TB"
    sStr = sStr & "   WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND STDCD = '" & Trim(sStdCD) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            sRet = "":
            If IsNull(.Fields("STDNM")) = False Then
                sRet = Trim(.Fields("STDNM"))
            Else
                sRet = ""
            End If
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Find_StdNM = sRet
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�л��� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
    
    Find_StdNM = ""
    
End Function
    
Private Sub cmdFind_Click()
    Call inits
    Call Find_Monthly                   '<< �����ڷ� ��ȸ
    
    If nTotRec > 0 Then
        Call Disp_STD_JumsuData(1)      '<< �л����� �ڷ� disp
        
    Else
        MsgBox "���輺�� ������ �����ϴ�.", vbExclamation + vbOKOnly, "�л����� ������"
    End If
    
    
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
            VScroll1.Value = nS - 1
            VScroll1.Enabled = False
                Call Disp_STD_JumsuData(VScroll1.Value)
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
            VScroll1.Value = nS + 1
            VScroll1.Enabled = False
                Call Disp_STD_JumsuData(VScroll1.Value)
            VScroll1.Enabled = True
        End If
    End If
End Sub

'>> scroll �̵�
Private Sub VScroll1_Change()
    If Me.Tag = "LOAD" Then Exit Sub
    
    VScroll1.Enabled = False
        Call Disp_STD_JumsuData(VScroll1.Value)
        txtPage.Text = Trim(CStr(VScroll1.Value)) & "/" & Trim(CStr(nTotRec))
    VScroll1.Enabled = True
End Sub


'## �л� ��������
Private Sub Disp_STD_JumsuData(ByVal aindex As Long)
    
    Dim nDay        As Integer
    Dim nPosition   As Integer
    
    Dim sBoldYM     As String
    Dim bBold       As Boolean
    Dim nDayText    As Integer
    
    If Me.Tag = "LOAD" Then Exit Sub
    If UBound(uSTD) < 1 Then Exit Sub
    
    sBoldYM = Format(Now, "yyyy-mm")            '< ����� ����
    
    Select Case uSTD(aindex).GAEYOL
        Case "01"
            txtGaeyol.Text = "�ι�"
        Case "02"
            txtGaeyol.Text = "�ڿ�"
        Case "03"
            txtGaeyol.Text = "��ü"
    End Select
    
    txtBan.Text = uSTD(aindex).BAN
    txtStdCD1.Text = Right(uSTD(aindex).STDCD, 4)
    txtStdNM1.Text = uSTD(aindex).STDNM
    
    bBold = False
    M1D(0).Text = uSTD(aindex).M1D(0)
    If StrComp(M1D(0).Text, sBoldYM, vbTextCompare) = 0 Then bBold = True
    M1D(0).FontSize = 9:        If bBold = True Then M1D(0).FontSize = 11
    M1D(0).FontBold = bBold
    
    '2011-05-19 ���ѿ� �뷮�� �뼺 �����쾾 ��û���� �����Ͽ� �ش��ϴ� �����͸� �߾ӿ� ǥ��
    
    nPosition = 3
    
    For nDay = 1 To 31 Step 1
        For nDayText = nPosition To 23 Step 5
            If Mid(uSTD(aindex).M1D(nDay), 7, 10) = "(��)" Then
                M1D(nDayText).Text = uSTD(aindex).M1D(nDay):    M1D(nDayText).FontBold = bBold
                                                            M1D(nDayText).FontSize = 9:     If bBold = True Then M1D(nDayText).FontSize = 11
                M1N(nDayText).Text = uSTD(aindex).M1N(nDay):    M1N(nDayText).FontBold = bBold
                                                            M1N(nDayText).FontSize = 9:     If bBold = True Then M1N(nDayText).FontSize = 11
                nPosition = nPosition + 5
                Exit For
            End If
        Next nDayText
    Next nDay
    
    bBold = False
    M2D(0).Text = uSTD(aindex).M2D(0)
    If StrComp(M2D(0).Text, sBoldYM, vbTextCompare) = 0 Then bBold = True
    M2D(0).FontSize = 9:        If bBold = True Then M2D(0).FontSize = 11
    M2D(0).FontBold = bBold
    
    nPosition = 3
    
    For nDay = 1 To 31 Step 1
        For nDayText = nPosition To 23 Step 5
            If Mid(uSTD(aindex).M2D(nDay), 7, 10) = "(��)" Then
                M2D(nDayText).Text = uSTD(aindex).M2D(nDay):    M2D(nDayText).FontBold = bBold
                                                            M2D(nDayText).FontSize = 9:     If bBold = True Then M2D(nDayText).FontSize = 11
                M2N(nDayText).Text = uSTD(aindex).M2N(nDay):    M2N(nDay).FontBold = bBold
                                                            M2N(nDayText).FontSize = 9:     If bBold = True Then M2N(nDayText).FontSize = 11
                nPosition = nPosition + 5
                Exit For
            End If
        Next nDayText
    Next nDay
    
    bBold = False
    M3D(0).Text = uSTD(aindex).M3D(0)
    If StrComp(M3D(0).Text, sBoldYM, vbTextCompare) = 0 Then bBold = True
    M3D(0).FontSize = 9:        If bBold = True Then M3D(0).FontSize = 11
    M3D(0).FontBold = bBold
    
    nPosition = 3
    
    For nDay = 1 To 31 Step 1
        For nDayText = nPosition To 23 Step 5
            If Mid(uSTD(aindex).M3D(nDay), 7, 10) = "(��)" Then
                M3D(nDayText).Text = uSTD(aindex).M3D(nDay):    M3D(nDayText).FontBold = bBold
                                                            M3D(nDayText).FontSize = 9:     If bBold = True Then M3D(nDayText).FontSize = 11
                M3N(nDayText).Text = uSTD(aindex).M3N(nDay):    M3N(nDayText).FontBold = bBold
                                                            M3N(nDayText).FontSize = 9:     If bBold = True Then M3N(nDayText).FontSize = 11
                nPosition = nPosition + 5
                Exit For
            End If
        Next nDayText
    Next nDay
    
    bBold = False
    M4D(0).Text = uSTD(aindex).M4D(0)
    If StrComp(M4D(0).Text, sBoldYM, vbTextCompare) = 0 Then bBold = True
    M4D(0).FontSize = 9:        If bBold = True Then M4D(0).FontSize = 11
    M4D(0).FontBold = bBold
    
    nPosition = 3
    
    For nDay = 1 To 31 Step 1
        For nDayText = nPosition To 23 Step 5
            If Mid(uSTD(aindex).M4D(nDay), 7, 10) = "(��)" Then
                M4D(nDayText).Text = uSTD(aindex).M4D(nDay):    M4D(nDayText).FontBold = bBold
                                                            M4D(nDayText).FontSize = 9:     If bBold = True Then M4D(nDayText).FontSize = 11
                M4N(nDayText).Text = uSTD(aindex).M4N(nDay):    M4N(nDayText).FontBold = bBold
                                                            M4N(nDayText).FontSize = 9:     If bBold = True Then M4N(nDayText).FontSize = 11
                nPosition = nPosition + 5
                Exit For
            End If
        Next nDayText
    Next nDay
    
    bBold = False
    M5D(0).Text = uSTD(aindex).M5D(0)
    If StrComp(M5D(0).Text, sBoldYM, vbTextCompare) = 0 Then bBold = True
    M5D(0).FontSize = 9:        If bBold = True Then M5D(0).FontSize = 11
    M5D(0).FontBold = bBold
    
    nPosition = 3
    
    For nDay = 1 To 31 Step 1
        For nDayText = nPosition To 23 Step 5
            If Mid(uSTD(aindex).M5D(nDay), 7, 10) = "(��)" Then
                M5D(nDayText).Text = uSTD(aindex).M5D(nDay):    M5D(nDayText).FontBold = bBold
                                                            M5D(nDayText).FontSize = 9:     If bBold = True Then M5D(nDayText).FontSize = 11
                M5N(nDayText).Text = uSTD(aindex).M5N(nDay):    M5N(nDayText).FontBold = bBold
                                                            M5N(nDayText).FontSize = 9:     If bBold = True Then M5N(nDayText).FontSize = 11
                nPosition = nPosition + 5
                Exit For
            End If
        Next nDayText
    Next nDay
    
End Sub


'## �л� ������ȸ (all)
Private Sub Find_Monthly()
    Dim sLastDay    As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim sStr        As String
    Dim nDay        As Integer
    Dim sTmp        As String
    Dim sFieldNM    As String
    
    Dim nRec        As Long
    
    Dim sYM1        As String
    Dim sYMD1       As String
    
    Dim sYM2        As String
    Dim sYMD2       As String
    
    Dim sYM3        As String
    Dim sYMD3       As String
    
    Dim sYM4        As String
    Dim sYMD4       As String
    
    Dim sYM5        As String
    Dim sYMD5       As String
    
    Dim sDayNM      As String
    
    Me.Tag = "LOAD"
    progDisp.Visible = True
    nTotRec = 0
    
    On Error GoTo ErrStmt
    
    sYM1 = Left(fpExmYM.Text, 7)
    sYMD1 = Format(DateAdd("m", 1, CDate(sYM1 & "-01")) - 1, "DD")
    
    sYM2 = Format(DateAdd("m", 1, CDate(Left(fpExmYM.Text, 7) & "-01")), "YYYY-MM")
    sYMD2 = Format(DateAdd("m", 1, CDate(sYM2 & "-01")) - 1, "DD")
    
    sYM3 = Format(DateAdd("m", 2, CDate(Left(fpExmYM.Text, 7) & "-01")), "YYYY-MM")
    sYMD3 = Format(DateAdd("m", 1, CDate(sYM3 & "-01")) - 1, "DD")
    
    sYM4 = Format(DateAdd("m", 3, CDate(Left(fpExmYM.Text, 7) & "-01")), "YYYY-MM")
    sYMD4 = Format(DateAdd("m", 1, CDate(sYM4 & "-01")) - 1, "DD")
    
    sYM5 = Format(DateAdd("m", 4, CDate(Left(fpExmYM.Text, 7) & "-01")), "YYYY-MM")
    sYMD5 = Format(DateAdd("m", 1, CDate(sYM5 & "-01")) - 1, "DD")
    
    
    sStr = ""
    sStr = sStr & " SELECT ACID, STDCD, MAX(STDNM) AS STDNM, MAX(GAEYOL) AS GAEYOL, MAX(BAN) AS BAN, "
    sStr = sStr & "        MAX(A) AS A, MAX(B) AS B, "
    
    For nDay = 1 To CInt(sYMD1) Step 1
        sStr = sStr & "    MAX( D1" & Format(nDay, "00") & " ) AS D1" & Format(nDay, "00") & ","
    Next nDay
    
    For nDay = 1 To CInt(sYMD2) Step 1
        sStr = sStr & "    MAX( D2" & Format(nDay, "00") & " ) AS D2" & Format(nDay, "00") & ","
    Next nDay
    
    For nDay = 1 To CInt(sYMD3) Step 1
        sStr = sStr & "    MAX( D3" & Format(nDay, "00") & " ) AS D3" & Format(nDay, "00") & ","
    Next nDay
    
    For nDay = 1 To CInt(sYMD4) Step 1
        sStr = sStr & "    MAX( D4" & Format(nDay, "00") & " ) AS D4" & Format(nDay, "00") & ","
    Next nDay
    
    For nDay = 1 To CInt(sYMD5) Step 1
        sStr = sStr & "    MAX( D5" & Format(nDay, "00") & " ) AS D5" & Format(nDay, "00") & ","
    Next nDay
    
    sStr = sStr & "        MAX(REGDAY) AS REGDAY "
    sStr = sStr & "   FROM ("
    
            sStr = sStr & " SELECT A.ACID, A.STDCD, A.STDNM, A.GAEYOL, A.BAN, '' AS A, '' AS B,"
        '>> 1 ��
            For nDay = 1 To CInt(sYMD1) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Replace(sYM1, "-", "", 1, -1, vbTextCompare) & Format(nDay, "00") & "', E_NUM) AS D1" & Format(nDay, "00") & ","
            Next nDay
            
        '>> 2 ��
            For nDay = 1 To CInt(sYMD2) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Replace(sYM2, "-", "", 1, -1, vbTextCompare) & Format(nDay, "00") & "', E_NUM) AS D2" & Format(nDay, "00") & ","
            Next nDay
            
        '>> 3 ��
            For nDay = 1 To CInt(sYMD3) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Replace(sYM3, "-", "", 1, -1, vbTextCompare) & Format(nDay, "00") & "', E_NUM) AS D3" & Format(nDay, "00") & ","
            Next nDay
            
        '>> 4 ��
            For nDay = 1 To CInt(sYMD4) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Replace(sYM4, "-", "", 1, -1, vbTextCompare) & Format(nDay, "00") & "', E_NUM) AS D4" & Format(nDay, "00") & ","
            Next nDay
            
        '>> 5 ��
            For nDay = 1 To CInt(sYMD5) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Replace(sYM5, "-", "", 1, -1, vbTextCompare) & Format(nDay, "00") & "', E_NUM) AS D5" & Format(nDay, "00") & ","
            Next nDay
            
            sStr = sStr & "        A.REGDAY "
            sStr = sStr & "   FROM SDEXM10TB A, SDEXM11TB B"
            sStr = sStr & "  Where A.STDCD  = B.STDCD"
            sStr = sStr & "    AND B.EXMDAY BETWEEN '" & Replace(sYM1, "-", "", 1, -1, vbTextCompare) & "01'"
            sStr = sStr & "                     AND '" & Replace(sYM5, "-", "", 1, -1, vbTextCompare) & sYMD5 & "'"
            
            If Trim(fpSTD_Ns.UnFmtText) > " " And Trim(fpSTD_Ne.UnFmtText) > " " Then
                sStr = sStr & " AND STDCD   BETWEEN '" & Trim(fpSTD_Ns.UnFmtText) & "'"
                sStr = sStr & "                 AND '" & Trim(fpSTD_Ne.UnFmtText) & "'"
            End If
            
            If Trim(Right(cboKaeyol.Text, 10)) <> "ALL" Then
                sStr = sStr & "        AND A.GAEYOL = '" & Trim(Right(cboKaeyol.Text, 5)) & "'"
            End If
            If Trim(Right(cboBan.Text, 4)) <> "ALL" Then
                sStr = sStr & "        AND A.BAN    = '" & Trim(Right(cboBan.Text, 5)) & "'"
            End If
            
            If Trim(txtStdCD.Text) > " " Then
                sStr = sStr & " AND STDCD = '" & Trim(txtStdCD.Text) & "'"
            End If
            If Trim(txtStdNM.Text) > " " Then
                sStr = sStr & " AND STDNM = '" & Trim(txtStdNM.Text) & "'"
            End If
            
    sStr = sStr & "          )"
    sStr = sStr & "  GROUP BY ACID, STDCD "
    sStr = sStr & "  ORDER BY ACID, STDCD "


    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec

        If .RecordCount > 0 Then
            .MoveFirst
            
            ReDim uSTD(.RecordCount) As tSTD                '<< �л��� ��ŭ ���ڵ� ����.
            
            nTotRec = .RecordCount
            
            VScroll1.Max = .RecordCount
            VScroll1.Enabled = True
            VScroll1.Value = 1
            
            txtPage.Text = "1/" & Trim(CStr(nTotRec))
            
        End If

        For nRec = 1 To .RecordCount Step 1

            progDisp.Value = Format(nRec / .RecordCount * 100, "##0")

            sTmp = "":      If IsNull(.Fields("ACID")) = False Then sTmp = Trim(.Fields("ACID")):                   uSTD(nRec).ACID = sTmp
            sTmp = "":      If IsNull(.Fields("STDCD")) = False Then sTmp = Trim(.Fields("STDCD")):                 uSTD(nRec).STDCD = sTmp
            sTmp = "":      If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM")):                 uSTD(nRec).STDNM = sTmp
            sTmp = "":      If IsNull(.Fields("GAEYOL")) = False Then sTmp = Trim(.Fields("GAEYOL")):               uSTD(nRec).GAEYOL = sTmp
            sTmp = "":      If IsNull(.Fields("BAN")) = False Then sTmp = Trim(.Fields("BAN")):                     uSTD(nRec).BAN = sTmp
            
            uSTD(nRec).M1D(0) = sYM1
            For nDay = 1 To CInt(sYMD1) Step 1
                sFieldNM = "D1" & Format(nDay, "00")
                sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                
                Select Case Weekday(CDate(sYM1 & "-" & Format(nDay, "00")))
                    Case 1
                        sDayNM = " (��)"
                    Case 2
                        sDayNM = " (��)"
                    Case 3
                        sDayNM = " (ȭ)"
                    Case 4
                        sDayNM = " (��)"
                    Case 5
                        sDayNM = " (��)"
                    Case 6
                        sDayNM = " (��)"
                    Case 7
                        sDayNM = " (��)"
                End Select
                
                If sDayNM = " (��)" Then
                uSTD(nRec).M1D(nDay) = Mid(sYM1, 6, 2) & "/" & Format(nDay, "00") & sDayNM
                uSTD(nRec).M1N(nDay) = sTmp
                End If
            Next nDay
            
            uSTD(nRec).M2D(0) = sYM2
            For nDay = 1 To CInt(sYMD2) Step 1
                sFieldNM = "D2" & Format(nDay, "00")
                sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                
                Select Case Weekday(CDate(sYM2 & "-" & Format(nDay, "00")))
                    Case 1
                        sDayNM = " (��)"
                    Case 2
                        sDayNM = " (��)"
                    Case 3
                        sDayNM = " (ȭ)"
                    Case 4
                        sDayNM = " (��)"
                    Case 5
                        sDayNM = " (��)"
                    Case 6
                        sDayNM = " (��)"
                    Case 7
                        sDayNM = " (��)"
                End Select
                If sDayNM = " (��)" Then
                uSTD(nRec).M2D(nDay) = Mid(sYM2, 6, 2) & "/" & Format(nDay, "00") & sDayNM
                uSTD(nRec).M2N(nDay) = sTmp
                End If
            Next nDay
            
            uSTD(nRec).M3D(0) = sYM3
            For nDay = 1 To CInt(sYMD3) Step 1
                sFieldNM = "D3" & Format(nDay, "00")
                sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                
                Select Case Weekday(CDate(sYM3 & "-" & Format(nDay, "00")))
                    Case 1
                        sDayNM = " (��)"
                    Case 2
                        sDayNM = " (��)"
                    Case 3
                        sDayNM = " (ȭ)"
                    Case 4
                        sDayNM = " (��)"
                    Case 5
                        sDayNM = " (��)"
                    Case 6
                        sDayNM = " (��)"
                    Case 7
                        sDayNM = " (��)"
                End Select
                If sDayNM = " (��)" Then
                uSTD(nRec).M3D(nDay) = Mid(sYM3, 6, 2) & "/" & Format(nDay, "00") & sDayNM
                uSTD(nRec).M3N(nDay) = sTmp
                End If
            Next nDay
            
            uSTD(nRec).M4D(0) = sYM4
            For nDay = 1 To CInt(sYMD4) Step 1
                sFieldNM = "D4" & Format(nDay, "00")
                sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                
                Select Case Weekday(CDate(sYM4 & "-" & Format(nDay, "00")))
                    Case 1
                        sDayNM = " (��)"
                    Case 2
                        sDayNM = " (��)"
                    Case 3
                        sDayNM = " (ȭ)"
                    Case 4
                        sDayNM = " (��)"
                    Case 5
                        sDayNM = " (��)"
                    Case 6
                        sDayNM = " (��)"
                    Case 7
                        sDayNM = " (��)"
                End Select
                If sDayNM = " (��)" Then
                uSTD(nRec).M4D(nDay) = Mid(sYM4, 6, 2) & "/" & Format(nDay, "00") & sDayNM
                uSTD(nRec).M4N(nDay) = sTmp
                End If
            Next nDay
            
            uSTD(nRec).M5D(0) = sYM5
            For nDay = 1 To CInt(sYMD5) Step 1
                sFieldNM = "D5" & Format(nDay, "00")
                sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                
                Select Case Weekday(CDate(sYM5 & "-" & Format(nDay, "00")))
                    Case 1
                        sDayNM = " (��)"
                    Case 2
                        sDayNM = " (��)"
                    Case 3
                        sDayNM = " (ȭ)"
                    Case 4
                        sDayNM = " (��)"
                    Case 5
                        sDayNM = " (��)"
                    Case 6
                        sDayNM = " (��)"
                    Case 7
                        sDayNM = " (��)"
                End Select
                If sDayNM = " (��)" Then
                uSTD(nRec).M5D(nDay) = Mid(sYM5, 6, 2) & "/" & Format(nDay, "00") & sDayNM
                uSTD(nRec).M5N(nDay) = sTmp
                End If
            Next nDay

            .MoveNext

        Next nRec
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing

    progDisp.Visible = False
    
    Me.Tag = ""

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    progDisp.Visible = False
    Me.Tag = ""

    On Error GoTo 0
    MsgBox "�л� ������ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption

End Sub







'## ��ü ���
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
        MsgBox "�μ�����մϴ�.", vbExclamation + vbOKOnly, "���������� �μ��ϱ�"
        Exit Sub
    End If
    
    nRec = 0
    cmdPrint.Tag = "ALL"
    
    Do
        nRec = nRec + 1
        txtPage.Text = Trim(CStr(nRec)) & "/" & Trim(CStr(UBound(uSTD)))
        
        
        Call Disp_STD_JumsuData(nRec)                           '<< �л��ڷ� ȭ�� ���̱�
        Me.Tag = "LOAD"
            VScroll1.Value = nRec
            Call CmdPrint_Click:        DoEvents                '<< 1�� ���
            
        Me.Tag = ""

    Loop Until nRec = UBound(uSTD)
    
    cmdPrint.Tag = ""
    MsgBox "����� �Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ü���"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    cmdPrint.Tag = ""
    
    MsgBox "��½� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "��ü���"
    
End Sub

'## ���� ������ ��� : 1�� ���
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
    
    '<< ���� �������� ����ϸ�,
    If cmdPrint.Tag = "" Then
        bChk = False
        With dlgPrint
            .CancelError = True
            .ShowPrinter
            
            bChk = True
        End With
        
ErrPrint:
        If bChk = False Then
            MsgBox "�μ�����մϴ�.", vbExclamation + vbOKOnly, "���������� �μ��ϱ�"
            Exit Sub
        End If
    End If
    
    '****************************************************************************************
    ' ������ ����ʱ�ȭ�� �Ѵ�.
    ' PrintStartDoc (Width, Height, PaperSize, Orientation,TopMargin,LeftMargin
    '****************************************************************************************
    pRate = 1.15
    basFunction.PrintStartDoc pReportViewer.Width * pRate, pReportViewer.Height * pRate, vbPRPSA4, vbPRORLandscape, 1, 1


    '********************************************************************
    '  �÷����� �̿��Ͽ� CONTROL�� �迭�� ó���Ѵ�.
    ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '  �� �Ʒ��� ������ ����� �ٲ��� ����....   boss
    '********************************************************************
    Dim UsrCtl      As Control

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS") Then
                If .Visible = True Then
                    '********************************************************************
                    '  �׵θ� ���� �簢 �ڽ��� ����� ���λ��� ĥ�Ѵ�.
                    '********************************************************************
                     Printer.DrawWidth = 1                   ' ���� ����
                     Printer.FillStyle = vbFSTransparent     ' �ܻ�
                     Printer.FillColor = &HC1F1FF            ' ���� ĥ�ϱ�
                     PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
                End If
             End If
        End With
    Next

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS") Then
                If .Visible = True Then
                    '********************************************************************
                    '  line�� �̿��� box�����(�⺻������ shape�� ��½� line�� �̿��Ѵ�)
                    '********************************************************************
                     Printer.DrawWidth = 12
                     PrintBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
                End If
             End If
        End With
    Next


    For Each UsrCtl In Me
        With UsrCtl
             Select Case UCase(TypeName(UsrCtl))
                    Case "LINE"
                        If .Visible = True Then
                            '********************************************************************
                            '  �ڽ�/line�� �ߴ´�.
                            '********************************************************************
                             Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                             Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                             Printer.FillStyle = vbFSTransparent
                             PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate
                        End If
                    Case "LABEL"
                        If .Visible = True Then
                            '********************************************************************
                            '  Label�� �״�� ��� �Ѵ�(�Ӽ�)
                            '  ��) transparent�� true�� ó���ϰ� �����Ѵ�.
                            '  SetBkMode(Printer.hdc, TRANSPARENT)������ MS���׸� ó���ϱ� ����
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
                        End If

                    Case "TEXTBOX"
                        If .Visible = True Then
                            If .Text = "" Or .Text = "0" Then
                                'no action
                            Else
                                '********************************************************************
                                '  ������ ��� (DATA�� TEXTBOX�� ó�� �Ѵ�.)
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
                            End If
                        End If
                    Case "IMAGE"
                        If .Visible = True Then
                            '********************************************************************
                            '  �������
                            '********************************************************************
                            If (Photo.Picture <> 0) Then
                                Printer.FontTransparent = True
                                iBKMode = SetBkMode(Printer.hDC, OPAQUE)
                                ' iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
                                PrintPicture .Picture, .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
                            End If
                        End If
             End Select
        End With
    Next

    Printer.EndDoc     ' �����ͷ� ������

End Sub


















