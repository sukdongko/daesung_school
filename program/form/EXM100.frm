VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form EXM100 
   Caption         =   "�л��������"
   ClientHeight    =   11640
   ClientLeft      =   1800
   ClientTop       =   3165
   ClientWidth     =   19080
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11640
   ScaleWidth      =   19080
   Begin VB.Frame fraStdin 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   8955
      Left            =   18360
      TabIndex        =   36
      Top             =   390
      Width           =   5895
      Begin VB.Frame Frame6 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Height          =   945
         Left            =   30
         TabIndex        =   39
         Top             =   30
         Width           =   5835
         Begin VB.CommandButton cmdStdDel 
            Caption         =   "�л������ϱ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3600
            TabIndex        =   51
            Top             =   510
            Width           =   2145
         End
         Begin VB.TextBox txtStdCDin 
            Height          =   345
            Left            =   540
            TabIndex        =   2
            Text            =   "txtStdCDin"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtStdNMin 
            Height          =   345
            Left            =   1440
            TabIndex        =   3
            Text            =   "txtStdNMin"
            Top             =   480
            Width           =   1005
         End
         Begin VB.TextBox txtBan 
            Height          =   345
            Left            =   1950
            TabIndex        =   1
            Text            =   "txtBan"
            Top             =   90
            Width           =   525
         End
         Begin VB.ComboBox cboGaeyol 
            Height          =   300
            Left            =   540
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   0
            Top             =   90
            Width           =   915
         End
         Begin VB.CommandButton cmdStdin 
            Caption         =   "�л�����ϱ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3600
            TabIndex        =   4
            Top             =   90
            Width           =   2145
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�л�"
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
            Left            =   120
            TabIndex        =   42
            Top             =   555
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "��"
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
            Index           =   0
            Left            =   1710
            TabIndex        =   41
            Top             =   180
            Width           =   615
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
            Left            =   120
            TabIndex        =   40
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Height          =   7905
         Left            =   30
         TabIndex        =   37
         Top             =   1020
         Width           =   5835
         Begin VB.CommandButton cmdAllSTD 
            Caption         =   "�л��ϰ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3630
            TabIndex        =   5
            Top             =   60
            Width           =   2115
         End
         Begin FPSpread.vaSpread sprStdin 
            Height          =   7035
            Left            =   90
            TabIndex        =   38
            Top             =   810
            Width           =   5655
            _Version        =   393216
            _ExtentX        =   9975
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
            MaxCols         =   5
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "EXM100.frx":0000
         End
         Begin MSComctlLib.ProgressBar progStdin 
            Height          =   135
            Left            =   90
            TabIndex        =   43
            Top             =   750
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin EditLib.fpLongInteger fpStdinRow 
            Height          =   285
            Left            =   4740
            TabIndex        =   6
            Top             =   480
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
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
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "���߰��� +, ������ - Ű�� �����ð� �����Ͻʽÿ�."
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
            Left            =   180
            TabIndex        =   44
            Top             =   570
            Width           =   4575
         End
      End
   End
   Begin VB.Frame fraAllJumsuin 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   9285
      Left            =   12540
      TabIndex        =   32
      Top             =   3300
      Width           =   5895
      Begin VB.Frame Frame2 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Height          =   9225
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   5835
         Begin FPSpread.vaSpread sprAllSaves 
            Height          =   8325
            Left            =   90
            TabIndex        =   35
            Top             =   840
            Width           =   5655
            _Version        =   393216
            _ExtentX        =   9975
            _ExtentY        =   14684
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
            MaxCols         =   5
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "EXM100.frx":192C
         End
         Begin VB.CommandButton cmdAllSaves 
            Caption         =   "����ϱ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3390
            TabIndex        =   34
            Top             =   60
            Width           =   2145
         End
         Begin MSComctlLib.ProgressBar progAllSaves 
            Height          =   135
            Left            =   90
            TabIndex        =   46
            Top             =   780
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin EditLib.fpLongInteger fpAllSaves 
            Height          =   285
            Left            =   4740
            TabIndex        =   47
            Top             =   480
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
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
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "���߰��� +, ������ - Ű�� �����ð� �����Ͻʽÿ�."
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
            Index           =   4
            Left            =   210
            TabIndex        =   48
            Top             =   570
            Width           =   4575
         End
      End
   End
   Begin VB.Frame fraView1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   9285
      Left            =   60
      TabIndex        =   24
      Top             =   810
      Width           =   14865
      Begin VB.Frame fraView11 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Height          =   9225
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   14805
         Begin VB.CommandButton cmdExcel 
            Caption         =   "�ݺ��϶�ǥ ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   12120
            TabIndex        =   50
            Top             =   120
            Width           =   2145
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "���� �л����� ��ȸ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9780
            TabIndex        =   17
            Top             =   120
            Width           =   2145
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3660
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   15
            Top             =   150
            Width           =   1485
         End
         Begin VB.ComboBox cboBan 
            Height          =   300
            Left            =   6390
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   16
            Top             =   120
            Width           =   1485
         End
         Begin EditLib.fpDateTime fpExmYM 
            Height          =   330
            Left            =   1320
            TabIndex        =   14
            Top             =   120
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
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
         Begin FPSpread.vaSpread sprSTD 
            Height          =   8655
            Left            =   30
            TabIndex        =   18
            Top             =   540
            Width           =   14745
            _Version        =   393216
            _ExtentX        =   26009
            _ExtentY        =   15266
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
            MaxCols         =   40
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "EXM100.frx":3230
         End
         Begin MSComctlLib.ProgressBar progDisp 
            Height          =   135
            Left            =   30
            TabIndex        =   49
            Top             =   480
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label5 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����"
            Height          =   210
            Left            =   210
            TabIndex        =   30
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   2610
            TabIndex        =   27
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��"
            Height          =   210
            Left            =   5100
            TabIndex        =   26
            Top             =   180
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraGwamok 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   645
      Left            =   30
      TabIndex        =   19
      Top             =   120
      Width           =   14895
      Begin VB.Frame Frame23 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Height          =   585
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   14835
         Begin VB.CommandButton cmdStdinShow 
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5070
            TabIndex        =   45
            Top             =   60
            Width           =   795
         End
         Begin VB.CommandButton cmdAllSave 
            Caption         =   "�ϰ�����ϱ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12570
            TabIndex        =   31
            Top             =   60
            Width           =   1935
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "�л��������(&P)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10410
            TabIndex        =   13
            Top             =   60
            Width           =   1935
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   4020
            TabIndex        =   9
            Text            =   "txtStdNM"
            Top             =   90
            Width           =   1005
         End
         Begin VB.TextBox txtStdCD 
            Height          =   345
            Left            =   3180
            TabIndex        =   8
            Text            =   "txtStdCD"
            Top             =   90
            Width           =   855
         End
         Begin EditLib.fpDateTime fpExmDay 
            Height          =   345
            Left            =   930
            TabIndex        =   7
            Top             =   90
            Width           =   1605
            _Version        =   196608
            _ExtentX        =   2831
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
            Text            =   "2004-01-01"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "YYYY-MM-DD"
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
         Begin EditLib.fpLongInteger fpK_Num 
            Height          =   345
            Left            =   6510
            TabIndex        =   10
            Top             =   90
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
         Begin EditLib.fpLongInteger fpE_Num 
            Height          =   345
            Left            =   9390
            TabIndex        =   12
            Top             =   90
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
         Begin EditLib.fpLongInteger fpM_Num 
            Height          =   345
            Left            =   7830
            TabIndex        =   11
            Top             =   90
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
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
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л�"
            Height          =   210
            Left            =   2190
            TabIndex        =   29
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��������"
            Height          =   210
            Left            =   -90
            TabIndex        =   28
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label8 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   225
            Left            =   6840
            TabIndex        =   23
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ܱ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   8340
            TabIndex        =   22
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���"
            Height          =   225
            Left            =   5520
            TabIndex        =   21
            Top             =   180
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "EXM100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'## �л����� ��� �� ��ȸ

Option Explicit
















'## ���α׷� �ʱ�ȭ
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    
    '�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ� �л����� �Էºκ� �ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
    fpExmDay.Text = Format(Now, "yyyy-mm-dd")
    txtStdCD.Text = ""
    txtStdNM.Text = ""
    
    fpK_Num.Value = 0
    fpM_Num.Value = 0
    fpE_Num.Value = 0
    
    fpExmYM.Text = Format(Now, "yyyy-mm")
        
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
'                .AddItem "��.����ι�" & Space(30) & "07"
'                .AddItem "��.����ڿ�" & Space(30) & "08"
'
'                .AddItem "��ȭ�ι�" & Space(30) & "09"
'                .AddItem "��ȭ�ڿ�" & Space(30) & "10"
'
'            End If
            
            .ListIndex = 0
        End With
    
    sprAllSaves.MaxRows = 0
    fpAllSaves.Value = 0
    
    fraAllJumsuin.ZOrder 0
    fraAllJumsuin.Top = 780
    fraAllJumsuin.Left = 8670
    cmdAllSave.Caption = "�ϰ�����ϱ�"
    fraAllJumsuin.Visible = False
    
    progAllSaves.Max = 100
    progAllSaves.Min = 0
    progAllSaves.Value = 0
    progAllSaves.Visible = False
    
    
    '�ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ� �л���Ϻκ� �ƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢƢ�
    cboGaeyol.Clear
    cboGaeyol.AddItem "�ι�" & Space(10) & "01"
    cboGaeyol.AddItem "�ڿ�" & Space(10) & "02"
    cboGaeyol.ListIndex = 0
    
    txtBan.Text = ""
    txtStdCDin.Text = ""
    txtStdNMin.Text = ""
    
    progStdin.Max = 100
    progStdin.Min = 0
    progStdin.Value = 0
    progStdin.Visible = False
    
    sprStdin.MaxRows = 0
    fpStdinRow.Value = 0
    
    
    fraStdin.ZOrder 0
    fraStdin.Top = 600
    fraStdin.Left = 3240
    cmdStdinShow.Caption = "���"
    fraStdin.Visible = False
    
    
    '## �������
    Call CreateHeader
    
    progDisp.Min = 0
    progDisp.Max = 100
    progDisp.Value = 0
    progDisp.Visible = False
    
End Sub


Private Sub cmdAllSave_Click()
    If fraAllJumsuin.Visible = True Then
        fraAllJumsuin.Visible = False
        cmdAllSave.Caption = "�ϰ�����ϱ�"
        
    Else
        fraAllJumsuin.Visible = True
        cmdAllSave.Caption = "�ϰ���ϴݱ�"
        
    End If
End Sub



Private Sub cmdStdinShow_Click()
    If fraStdin.Visible = True Then
        fraStdin.Visible = False
        cmdStdinShow.Caption = "���"
        
    Else
        fraStdin.Visible = True
        cmdStdinShow.Caption = "�ݱ�"
        
    End If
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



'## ȭ�� ������ ����
Private Sub Form_Resize()
    
    fraView1.Width = Me.Width - 240
    fraView11.Width = fraView1.Width - 80
    fraView1.Height = Me.Height - 1500
    fraView11.Height = fraView1.Height - 80
    sprSTD.Width = fraView11.Width - 80
    sprSTD.Height = fraView11.Height - 600
End Sub





Private Sub fpExmYM_LostFocus()
    Dim sRet        As String
    
    sRet = CreateHeader
    
End Sub

Private Function CreateHeader() As String
    Dim sYM         As String
    Dim sLastDay    As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nCol        As Integer
    
    With sprSTD
        .MaxRows = 0
        .MaxCols = 40
        
        .Row = SpreadHeader
        .Col = 1:               .Text = "�л�No."
        .Col = .Col + 1:        .Text = "�л���"
        .Col = .Col + 1:        .Text = "��"
        .Col = .Col + 1:        .Text = ""
        .Col = .Col + 1:        .Text = ""
        
        sYM = Left(fpExmYM.Text, 7) & "-01"
        sLastDay = Format(DateAdd("m", 1, CDate(Left(fpExmYM.Text, 7) & "-01")) - 1, "DD")
        
        .Col = 5        'col ������ ����Ʈ
        For nCol = 1 To CInt(sLastDay) Step 1
            
            .Row = SpreadHeader
            .Col = .Col + 1
                .Text = Mid(fpExmYM.UnFmtText, 5, 2) & "." & Format(nCol, "00")
        Next nCol
        
        .MaxCols = CLng(sLastDay) + 5
        
    End With
    
    CreateHeader = sLastDay

End Function









'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�
'�� �л����
'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�

Private Sub cmdStdin_Click()
    
    Dim DBCmd           As ADODB.Command
    Dim sStr            As String
    Dim nExe            As Long
    
    
    If Trim(txtBan.Text) = "" Then
        MsgBox "�� ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtStdCDin.Text) = "" Then
        MsgBox "�л���ȣ ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    If Trim(txtStdNMin.Text) = "" Then
        MsgBox "�л��̸� ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & " INSERT INTO SDEXM10TB (ACID, STDCD, STDNM, GAEYOL, BAN, REGDAY)"
    sStr = sStr & "                VALUES ("
    sStr = sStr & "                       '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                       '" & Format(CLng(txtStdCDin.Text), "0000") & "',"
    sStr = sStr & "                       '" & Trim(txtStdNMin.Text) & "',"
    sStr = sStr & "                       '" & Format(CInt(Right(cboGaeyol.Text, 10)), "00") & "',"
    sStr = sStr & "                       '" & Trim(txtBan.Text) & "',"
    sStr = sStr & "                       SYSDATE"
    sStr = sStr & "                        )"
    
    DBCmd.CommandType = adCmdText
    DBCmd.CommandText = sStr
    nExe = 0:       DBCmd.Execute nExe, , -1
    
    If nExe = 0 Then
        sStr = ""
        sStr = sStr & " UPDATE SDEXM10TB "
        sStr = sStr & "    SET STDNM  = '" & Trim(txtStdNMin.Text) & "',"
        sStr = sStr & "        GAEYOL = '" & Format(CInt(Right(cboGaeyol.Text, 10)), "00") & "',"
        sStr = sStr & "        BAN    = '" & Trim(txtBan.Text) & "',"
        sStr = sStr & "        REGDAY = SYSDATE "
        sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "    AND STDCD  = '" & Format(CLng(txtStdCDin.Text), "0000") & "'"
        
        DBCmd.CommandType = adCmdText
        DBCmd.CommandText = sStr
        nExe = 0:       DBCmd.Execute nExe, , -1
        
        If nExe = 0 Then
            MsgBox "�л� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
            On Error GoTo 0
            
            Set DBCmd = Nothing
            Exit Sub
        End If
        
    End If
    
    MsgBox "�л� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    Set DBCmd = Nothing
    
End Sub

Private Sub cmdStdDel_Click()
    Dim DBCmd           As ADODB.Command
    Dim sStr            As String
    Dim nExe            As Long
    
    If Trim(txtBan.Text) = "" Then
        MsgBox "�� ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtStdCDin.Text) = "" Then
        MsgBox "�л���ȣ ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    If Trim(txtStdNMin.Text) = "" Then
        MsgBox "�л��̸� ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    On Error Resume Next
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDEXM10TB "
    sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND STDCD  = '" & Format(CLng(txtStdCDin.Text), "0000") & "'"
        
    DBCmd.CommandType = adCmdText
    DBCmd.CommandText = sStr
    nExe = 0:       DBCmd.Execute nExe, , -1
        
    If nExe = 0 Then
        MsgBox "�л� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
        On Error GoTo 0
        
        Set DBCmd = Nothing
        Exit Sub
    End If
    
    MsgBox "�л� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    
    Set DBCmd = Nothing
    
End Sub

    
'## �л���� ��
Private Sub sprStdin_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        sprStdin.MaxRows = sprStdin.MaxRows + 1
        sprStdin.Row = 1
        
    ElseIf KeyCode = vbKeySubtract Then
        If sprStdin.MaxRows = 0 Then Exit Sub
        
        sprStdin.Row = sprStdin.MaxRows
        sprStdin.DeleteRows sprStdin.Row, 1
        sprStdin.MaxRows = sprStdin.MaxRows - 1
        
    End If
End Sub

Private Sub fpStdinRow_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        sprStdin.MaxRows = fpStdinRow.Value
        
    End If
End Sub

Private Sub fpStdinRow_LostFocus()
    sprStdin.MaxRows = fpStdinRow.Value
    
End Sub





'## �л��ϰ����
Private Sub cmdAllSTD_Click()
    
    Dim nRow            As Long

    Dim DBCmd           As ADODB.Command
    Dim sStr            As String
    Dim nExe            As Long
    
    Dim sStdCD          As String
    Dim sStdNM          As String
    Dim sGaeyol         As String
    Dim sBan            As String
    
    If sprStdin.MaxRows = 0 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    progStdin.Value = 0
    progStdin.Visible = True
    
    
    On Error Resume Next
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn
    
    For nRow = 1 To sprStdin.MaxRows Step 1
    
        progStdin.Value = Format(nRow / sprStdin.MaxRows * 100, "##0")
        
        sStdCD = ""
        sStdNM = ""
        sGaeyol = ""
        sBan = ""
        
        sprStdin.Row = nRow
        sprStdin.Col = 1:                   sStdCD = Format(CLng(sprStdin.Text), "0000")
        sprStdin.Col = sprStdin.Col + 1:    sStdNM = Trim(sprStdin.Text)
        sprStdin.Col = sprStdin.Col + 1:    sGaeyol = Format(CInt(sprStdin.Text), "00")
        sprStdin.Col = sprStdin.Col + 1:    sBan = Trim(sprStdin.Text)
        
        If sStdCD > " " Then
            sStr = ""
            sStr = sStr & " INSERT INTO SDEXM10TB (ACID, STDCD, STDNM, GAEYOL, BAN, REGDAY)"
            sStr = sStr & "                VALUES ("
            sStr = sStr & "                       '" & Trim(basModule.SchCD) & "',"
            sStr = sStr & "                       '" & sStdCD & "',"
            sStr = sStr & "                       '" & sStdNM & "',"
            sStr = sStr & "                       '" & sGaeyol & "',"
            sStr = sStr & "                       '" & sBan & "',"
            sStr = sStr & "                       SYSDATE"
            sStr = sStr & "                        )"
            
            DBCmd.CommandType = adCmdText
            DBCmd.CommandText = sStr
            nExe = 0:       DBCmd.Execute nExe, , -1
            
            If nExe = 0 Then
                sStr = ""
                sStr = sStr & " UPDATE SDEXM10TB "
                sStr = sStr & "    SET STDNM  = '" & sStdNM & "',"
                sStr = sStr & "        GAEYOL = '" & sGaeyol & "',"
                sStr = sStr & "        BAN    = '" & sBan & "',"
                sStr = sStr & "        REGDAY = SYSDATE "
                sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                sStr = sStr & "    AND STDCD  = '" & sStdCD & "'"
                
                DBCmd.CommandType = adCmdText
                DBCmd.CommandText = sStr
                nExe = 0:       DBCmd.Execute nExe, , -1
                
                If nExe = 0 Then
                    basDataBase.DBConn.RollbackTrans
                    
                    Set DBCmd = Nothing
                    
                    MsgBox "�л� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
                    On Error GoTo 0
                    
                    progStdin.Visible = False
                    Exit Sub
                End If
                
            End If
        End If
        
    Next nRow
    
    '## ��������
    basDataBase.DBConn.CommitTrans
    
    Set DBCmd = Nothing
    
    MsgBox "�л� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    progStdin.Visible = False
    
End Sub
    
    
    
'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�
'�� �л� �������
'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�

Private Sub sprAllSaves_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyAdd Then
        sprAllSaves.MaxRows = sprAllSaves.MaxRows + 1
        sprAllSaves.Row = 1
        
    ElseIf KeyCode = vbKeySubtract Then
        If sprAllSaves.MaxRows = 0 Then Exit Sub
        
        sprAllSaves.Row = sprAllSaves.MaxRows
        sprAllSaves.DeleteRows sprAllSaves.Row, 1
        sprAllSaves.MaxRows = sprAllSaves.MaxRows - 1
        
    End If
End Sub

Private Sub fpAllSaves_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        sprAllSaves.MaxRows = fpAllSaves.Value
        
    End If
End Sub

Private Sub fpAllSaves_LostFocus()
    sprAllSaves.MaxRows = fpAllSaves.Value
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



'## �л����� �Ѱ� ����ϱ�
Private Sub cmdSave_Click()
    
    Dim DBCmd           As ADODB.Command
    Dim sStr            As String
    Dim nExe            As Long
    
    Dim nKnum           As Long
    Dim nMnum           As Long
    Dim nEnum           As Long
    
    
    If Trim(txtStdCD.Text) = "" Then
        MsgBox "�л���ȣ ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    If Trim(txtStdNM.Text) = "" Then
        MsgBox "�л��̸� ����", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If fpK_Num.Text = "" Then
        fpK_Num.Value = 0
    End If
    nKnum = fpK_Num.Value
    
    If fpM_Num.Text = "" Then
        fpM_Num.Value = 0
    End If
    nMnum = fpM_Num.Value
    
    If fpE_Num.Text = "" Then
        fpE_Num.Value = 0
    End If
    nEnum = fpE_Num.Value
    
    
    On Error Resume Next
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & " INSERT INTO SDEXM11TB (ACID, STDCD, EXMDAY, K_NUM, M_NUM, E_NUM, REGDAY)"
    sStr = sStr & "                VALUES ("
    sStr = sStr & "                       '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                       '" & Format(CLng(txtStdCD.Text), "0000") & "',"
    sStr = sStr & "                       '" & Left(fpExmDay.UnFmtText, 8) & "',"
    sStr = sStr & "                        " & CStr(nKnum) & ","
    sStr = sStr & "                        " & CStr(nMnum) & ","
    sStr = sStr & "                        " & CStr(nEnum) & ","
    sStr = sStr & "                       SYSDATE"
    sStr = sStr & "                        )"
    
    DBCmd.CommandType = adCmdText
    DBCmd.CommandText = sStr
    nExe = 0:       DBCmd.Execute nExe, , -1
    
    If nExe = 0 Then
        sStr = ""
        sStr = sStr & " UPDATE SDEXM11TB "
        sStr = sStr & "    SET K_NUM  =  " & CStr(nKnum) & ","
        sStr = sStr & "        M_NUM  =  " & CStr(nMnum) & ","
        sStr = sStr & "        E_NUM  =  " & CStr(nEnum) & ","
        sStr = sStr & "        REGDAY = SYSDATE "
        sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "    AND STDCD  = '" & Format(CLng(txtStdCD.Text), "0000") & "'"
        sStr = sStr & "    AND EXMDAY = '" & Left(fpExmDay.UnFmtText, 8) & "'"
        
        DBCmd.CommandType = adCmdText
        DBCmd.CommandText = sStr
        nExe = 0:       DBCmd.Execute nExe, , -1
        
        If nExe = 0 Then
            MsgBox "�л� ������Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
            On Error GoTo 0
            
            Set DBCmd = Nothing
            Exit Sub
        End If
        
    End If
    
    MsgBox "�л� ������ ����Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    Set DBCmd = Nothing
    
End Sub

'��系�� ���
Private Sub cmdAllSaves_Click()
    
    Dim nRow            As Long

    Dim DBCmd           As ADODB.Command
    Dim sStr            As String
    Dim nExe            As Long
    
    Dim sStdCD          As String
    Dim sStdNM          As String
    Dim sExmDay         As String
    
    Dim sKnum           As String
    Dim sMnum           As String
    Dim sEnum           As String
    
    If sprAllSaves.MaxRows = 0 Then
        MsgBox "����� �ڷᰡ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    progAllSaves.Value = 0
    progAllSaves.Visible = True
    
    On Error Resume Next
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn
    
    For nRow = 1 To sprAllSaves.MaxRows Step 1
    
        progAllSaves.Value = Format(nRow / sprAllSaves.MaxRows * 100, "##0")
        
        sStdCD = ""
        sStdNM = ""
        sExmDay = ""
        sEnum = ""
        
        sprAllSaves.Row = nRow
        sprAllSaves.Col = 1:                        sStdCD = Format(CLng(sprAllSaves.Text), "0000")
        sprAllSaves.Col = sprAllSaves.Col + 1:      sStdNM = Trim(sprAllSaves.Text)
        sprAllSaves.Col = sprAllSaves.Col + 1:      sExmDay = Replace(sprAllSaves.Text, "-", "", 1, -1, vbTextCompare)
        sprAllSaves.Col = sprAllSaves.Col + 1:
            If Trim(sprAllSaves.Text) = "" Then
                sEnum = "0"
            Else
                sEnum = Trim(sprAllSaves.Text)
            End If
        
        If sStdCD > " " Then
            sStr = ""
            sStr = sStr & " INSERT INTO SDEXM11TB (ACID, STDCD, EXMDAY, "
            'sStr = sStr & "                        K_NUM, M_NUM, "
            sStr = sStr & "                        E_NUM, REGDAY)"
            sStr = sStr & "                VALUES ("
            sStr = sStr & "                       '" & Trim(basModule.SchCD) & "',"
            sStr = sStr & "                       '" & sStdCD & "',"
            sStr = sStr & "                       '" & sExmDay & "',"
            'sStr = sStr & "                        " & CStr(sKnum) & ","
            'sStr = sStr & "                        " & CStr(sMnum) & ","
            sStr = sStr & "                        " & CStr(sEnum) & ","
            sStr = sStr & "                       SYSDATE"
            sStr = sStr & "                        )"
            
            DBCmd.CommandType = adCmdText
            DBCmd.CommandText = sStr
            nExe = 0:       DBCmd.Execute nExe, , -1
            
            If nExe = 0 Then
                sStr = ""
                sStr = sStr & " UPDATE SDEXM11TB "
                sStr = sStr & "    SET "
                'sStr = sStr & "        K_NUM  =  " & CStr(sKnum) & ","
                'sStr = sStr & "        M_NUM  =  " & CStr(sMnum) & ","
                sStr = sStr & "        E_NUM  =  " & CStr(sEnum) & ","
                sStr = sStr & "        REGDAY = SYSDATE "
                sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                sStr = sStr & "    AND STDCD  = '" & sStdCD & "'"
                sStr = sStr & "    AND EXMDAY = '" & sExmDay & "'"
                
                DBCmd.CommandType = adCmdText
                DBCmd.CommandText = sStr
                nExe = 0:       DBCmd.Execute nExe, , -1
                
                If nExe = 0 Then
                    basDataBase.DBConn.RollbackTrans
                    
                    Set DBCmd = Nothing
                    
                    MsgBox "�л� ������Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
                    On Error GoTo 0
                    
                    progAllSaves.Visible = False
                    Exit Sub
                End If
                
            End If
        End If
        
    Next nRow
    
    '## ��������
    basDataBase.DBConn.CommitTrans
    
    Set DBCmd = Nothing
    
    MsgBox "�л� ��������Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    progAllSaves.Visible = False
End Sub



'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�
'�� �л� ������ȸ
'�̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢̢�

Private Sub cmdFind_Click()
    Dim sLastDay        As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim sStr        As String
    Dim nDay        As Integer
    Dim sTmp        As String
    Dim sFieldNM    As String
    
    Dim nRec        As Long
    
    progDisp.Visible = True
    sLastDay = CreateHeader             '< ��� ������ ���ÿ� ���������� ó��
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT STDCD, MAX(STDNM) AS STDNM, MAX(BAN) AS BAN, "
    sStr = sStr & "        MAX(A) AS A, MAX(B) AS B, "
    For nDay = 1 To CInt(sLastDay) Step 1
        sStr = sStr & "    MAX( D" & Format(nDay, "00") & " ) AS D" & Format(nDay, "00") & ","
    Next nDay
    sStr = sStr & "        MAX(REGDAY) AS REGDAY "
    sStr = sStr & "   FROM ("
            sStr = sStr & " SELECT A.STDCD, A.STDNM, A.BAN, '' AS A, '' AS B,"
            For nDay = 1 To CInt(sLastDay) Step 1
                sStr = sStr & "    DECODE(EXMDAY, '" & Left(fpExmYM.UnFmtText, 6) & Format(nDay, "00") & "', E_NUM) AS D" & Format(nDay, "00") & ","
            Next nDay
            sStr = sStr & "        A.REGDAY "
            sStr = sStr & "   FROM SDEXM10TB A, SDEXM11TB B"
            sStr = sStr & "  Where A.STDCD = B.STDCD(+)"
            sStr = sStr & "    AND B.EXMDAY BETWEEN '" & Left(fpExmYM.UnFmtText, 6) & "01'"
            sStr = sStr & "                     AND '" & Left(fpExmYM.UnFmtText, 6) & "31'"
    If Trim(Right(cboKaeyol.Text, 10)) <> "ALL" Then
        sStr = sStr & "        AND A.GAEYOL = '" & Trim(Right(cboKaeyol.Text, 5)) & "'"
    End If
    If Trim(Right(cboBan.Text, 4)) <> "ALL" Then
        sStr = sStr & "        AND A.BAN    = '" & Trim(Right(cboBan.Text, 5)) & "'"
    End If
    sStr = sStr & "         )"
    sStr = sStr & "  GROUP BY STDCD "
    sStr = sStr & "  ORDER BY STDCD "
    
    
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
    
        If .RecordCount > 0 Then .MoveFirst
    
        For nRec = 1 To .RecordCount Step 1
        
            progDisp.Value = Format(nRec / .RecordCount * 100, "##0")
        
            sprSTD.MaxRows = sprSTD.MaxRows + 1
            sprSTD.Row = sprSTD.MaxRows
            
            sprSTD.Col = 1:                         sTmp = "":      If IsNull(.Fields("STDCD")) = False Then sTmp = Trim(.Fields("STDCD")):         Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 60, sTmp)
            sprSTD.Col = sprSTD.Col + 1:            sTmp = "":      If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM")):         Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 60, sTmp)
            sprSTD.Col = sprSTD.Col + 1:            sTmp = "":      If IsNull(.Fields("BAN")) = False Then sTmp = Trim(.Fields("BAN")):             Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 60, sTmp)
            
            sprSTD.Col = sprSTD.Col + 1:            sTmp = "":      If IsNull(.Fields("A")) = False Then sTmp = Trim(.Fields("A")):                 Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 60, sTmp)
            sprSTD.Col = sprSTD.Col + 1:            sTmp = "":      If IsNull(.Fields("B")) = False Then sTmp = Trim(.Fields("B")):                 Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 60, sTmp)
            
            For nDay = 1 To CInt(sLastDay) Step 1
                sFieldNM = "D" & Format(nDay, "00")
                sprSTD.Col = sprSTD.Col + 1:        sTmp = "":      If IsNull(.Fields(sFieldNM)) = False Then sTmp = Trim(.Fields(sFieldNM))
                If IsNumeric(sTmp) = True Then
                    Call basFunction.Set_SprType_Numeric(sprSTD, 0, 0, 9999, "", CDbl(sTmp))
                End If
            Next nDay
            
            .MoveNext
            
        Next nRec
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    progDisp.Visible = False
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    progDisp.Visible = False
    
    On Error GoTo 0
    MsgBox "�л� ������ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
    
End Sub


Private Sub cmdExcel_Click()
    
    Call sprSTD.ExportToExcel("�ݺ��϶�ǥ", "�ݺ��϶�ǥ", "c:\temp\dd")
    
End Sub



































